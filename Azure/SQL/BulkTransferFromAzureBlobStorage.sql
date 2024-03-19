-- Create Master Key with randomally generated password 
CREATE MASTER KEY ENCRYPTION BY PASSWORD = 'uPeesQg34i89o!';

-- Create scoped credential pointing to SQL server managed identity (in this case SQL server is "sql-easparx-prod")
CREATE DATABASE SCOPED CREDENTIAL Managed_BlobAccess 
WITH IDENTITY = 'Managed Identity';

-- Create External DataSource by using credentials established above
CREATE EXTERNAL DATA SOURCE blobstoragesource
    WITH (
        TYPE = BLOB_STORAGE,
		LOCATION = 'https://dlreportingdataprod.blob.core.windows.net/ashf-test',
        CREDENTIAL = Managed_BlobAccess
    );

CREATE TABLE #Test (Line1 VARCHAR(50))
 
BULK INSERT #Test
FROM 'Test.txt'
WITH (DATA_SOURCE = 'blobstoragesource',
      FORMAT = 'CSV')
 
select * from #Test