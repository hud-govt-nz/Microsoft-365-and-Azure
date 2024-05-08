-- ADD USERS

-- FAW USER ADD ON DB
CREATE USER fawreporter WITH PASSWORD = 'e#bYHFbHT#2aQn'
ALTER ROLE db_datareader ADD MEMBER [fawreporter];

	-- Create 'groups' on Database - assumes security group exists in Azure AD
	CREATE USER [Data Reporting - SQLDB - Reader] FROM EXTERNAL PROVIDER WITH DEFAULT_SCHEMA=[dbo];
	CREATE USER [Data Reporting - SQLDB - Writer] FROM EXTERNAL PROVIDER WITH DEFAULT_SCHEMA=[dbo];

	-- Grant database roles to reader and writer groups
	ALTER ROLE db_datareader ADD MEMBER [Data Reporting - SQLDB - Reader];
	ALTER ROLE db_datareader ADD MEMBER [Data Reporting - SQLDB - Writer];
	ALTER ROLE db_datawriter ADD MEMBER [Data Reporting - SQLDB - Writer];
	ALTER ROLE db_ddladmin ADD MEMBER [Data Reporting - SQLDB - Writer];

	-- Grant additional permissions to writer group
	GRANT CREATE TABLE TO [Data Reporting - SQLDB - Writer]; 
	GRANT EXECUTE TO [Data Reporting - SQLDB - Writer]; 
	GRANT CONTROL TO [Data Reporting - SQLDB - Writer]; 
	GRANT CREATE PROCEDURE TO [Data Reporting - SQLDB - Writer]; 
	GRANT CREATE SCHEMA TO [Data Reporting - SQLDB - Writer]; 
	GRANT VIEW DATABASE STATE TO [Data Reporting - SQLDB - Writer];

-- REMOVE USERS

	-- Remove additional permissions from writer group
	REVOKE CREATE TABLE FROM [Data Reporting - SQLDB - Writer];
	REVOKE EXECUTE FROM [Data Reporting - SQLDB - Writer];
	REVOKE CONTROL FROM [Data Reporting - SQLDB - Writer];
	REVOKE CREATE PROCEDURE FROM [Data Reporting - SQLDB - Writer];
	REVOKE CREATE SCHEMA FROM [Data Reporting - SQLDB - Writer];
	REVOKE VIEW DATABASE STATE FROM [Data Reporting - SQLDB - Writer];


	-- Remove database roles to reader and writer groups
	ALTER ROLE db_datareader DROP MEMBER [Data Reporting - SQLDB - Reader];
	ALTER ROLE db_datareader DROP MEMBER [Data Reporting - SQLDB - Writer];
	ALTER ROLE db_datawriter DROP MEMBER [Data Reporting - SQLDB - Writer];
	ALTER ROLE db_ddladmin DROP MEMBER [Data Reporting - SQLDB - Writer];

	-- Remove 'groups' on Database - assumes security group exists in Azure AD
	DROP USER [Data Reporting - SQLDB - Reader];
	DROP USER [Data Reporting - SQLDB - Writer];

-- Database Role Memberships
SELECT 
    DP1.name AS DatabaseRoleName,   
    ISNULL(DP2.name, 'No members') AS DatabaseUserName,
    DP2.type_desc AS PrincipalType,
    'Role' AS MemberType,
    NULL AS PermissionType,
    NULL AS PermissionObject,
    NULL AS SchemaName
FROM 
    sys.database_role_members AS DRM  
    RIGHT OUTER JOIN sys.database_principals AS DP1  
        ON DRM.role_principal_id = DP1.principal_id  
    LEFT OUTER JOIN sys.database_principals AS DP2  
        ON DRM.member_principal_id = DP2.principal_id  
WHERE 
    DP1.type = 'R'

UNION ALL

-- Database Permissions
SELECT DISTINCT 
    pr.name AS DatabaseRoleName,   
    pr.name AS DatabaseUserName,
    pr.type_desc AS PrincipalType,
    'User' AS MemberType,
    pe.permission_name AS PermissionType,
    o.[name] AS PermissionObject,
    s.[name] AS SchemaName
FROM 
    sys.database_principals AS pr 
    JOIN sys.database_permissions AS pe 
        ON pe.grantee_principal_id = pr.principal_id
    LEFT JOIN sys.objects AS o 
        ON (o.object_id = pe.major_id)
    LEFT JOIN sys.schemas AS s
        ON o.schema_id = s.schema_id
ORDER BY 
    DatabaseRoleName, DatabaseUserName, PrincipalType, MemberType, PermissionType, PermissionObject, SchemaName;


-- Check user permissions

	WITH UsersAndRoles (principal_name, sid, type) AS 
(
    SELECT DISTINCT prin.name, prin.sid, prin.type 
    FROM sys.database_principals prin 
        INNER JOIN ( SELECT *
                     FROM sys.database_permissions
                     WHERE type = 'CO' 
                        AND state IN ('G', 'W')
        ) perm 
            ON perm.grantee_principal_id = prin.principal_id 
        WHERE prin.type IN ('S', 'X', 'R', 'E', 'G')
    UNION ALL
    SELECT 
        user_name(rls.member_principal_id), prin.sid, prin.type
    FROM 
        UsersAndRoles cte
        INNER JOIN sys.database_role_members rls
            ON user_name(rls.role_principal_id) = cte.principal_name
        INNER JOIN sys.database_principals prin
            ON rls.member_principal_id = prin.principal_id
        WHERE cte.type = 'R'
),
Users (database_user, sid) AS
(
    SELECT principal_name, sid
    FROM UsersAndRoles
    WHERE type IN ('S', 'X', 'E', 'G')
        AND principal_name != 'dbo'
)
SELECT DISTINCT database_user AS [User], sid AS [SID]
    FROM Users
    WHERE sid != 0x01