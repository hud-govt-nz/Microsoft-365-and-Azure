@{
    AllNodes = @(
        @{
            NodeName                    = "localhost"
            PSDscAllowPlainTextPassword = $true;
            PSDscAllowDomainUser        = $true;
            #region Parameters
            # Default Value Used to Ensure a Configuration Data File is Generated
            ServerNumber = "0"

        }
    )
    NonNodeData = @(
        @{
            # Tenant's default verified domain name
            OrganizationName = "mhud.onmicrosoft.com"

            # The Id or Name of the tenant to authenticate against
            TenantId = "mhud.onmicrosoft.com"

            # Azure AD Application Id for Authentication
            ApplicationId = "c574d912-050b-4acc-b0e6-63bed2c4c562"

            # Thumbprint of the certificate to use for authentication
            CertificateThumbprint = "789C2920A65EE8D5F4C3DDF4487B46B6D29293E4"

        }
    )
}
