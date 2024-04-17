[PS] C:\Windows\system32>Test-OAuthConnectivity -Service ews -TargetUri https://mail.acceptance.mfat.govt.nz/ews/exchange.asmx -Mailbox 'Ashley Forde' | fl *


Task        : Checking EWS API Call Under Oauth
Detail      : The configuration was last successfully loaded at 1/01/0001 12:00:00 a.m. UTC. This was 1063127577 minutes ago.
              The token cache is being cleared because "use cached token" was set to false.
              Exchange Outbound Oauth Log:
              Client request ID: bc0ee239-d119-4b63-aa28-4a7355adbd62
              Information:Using custom InternetWebProxy http://10.38.161.34:8080/.
              
              Exchange Response Details:
              HTTP response message: 
              Exception:
              System.Net.WebException: The remote server returned an error: (403) Forbidden.
                 at System.Net.HttpWebRequest.GetRequestStream(TransportContext& context)
                 at System.Net.HttpWebRequest.GetRequestStream()
                 at Microsoft.Exchange.Monitoring.TestOAuthConnectivityHelper.SendExchangeOAuthRequest(ADUser user, String orgDomain, Uri targetUri, String& 
              diagnosticMessage, Boolean appOnly, Boolean useCachedToken, Boolean reloadConfig)
              
ResultType  : Error
Identity    : Microsoft.Exchange.Security.OAuth.ValidationResultNodeId
IsValid     : True
ObjectState : New

[PS] C:\Windows\system32>Test-OAuthConnectivity -Service ews -TargetUri https://outlook.orangetst.mfat.net.nz/EWS/exchange.asmx -Mailbox 'Ashley Forde' | fl *


Task        : Checking EWS API Call Under Oauth
Detail      : The configuration was last successfully loaded at 1/01/0001 12:00:00 a.m. UTC. This was 1063127578 minutes ago.
              The token cache is being cleared because "use cached token" was set to false.
              Exchange Outbound Oauth Log:
              Client request ID: 3ee4d016-7723-41ed-851d-66f5d5224f7b
              Information:Using custom InternetWebProxy http://10.38.161.34:8080/.
              
              Exchange Response Details:
              HTTP response message: 
              Exception:
              System.Net.WebException: The remote server returned an error: (403) Forbidden.
                 at System.Net.HttpWebRequest.GetRequestStream(TransportContext& context)
                 at System.Net.HttpWebRequest.GetRequestStream()
                 at Microsoft.Exchange.Monitoring.TestOAuthConnectivityHelper.SendExchangeOAuthRequest(ADUser user, String orgDomain, Uri targetUri, String& 
              diagnosticMessage, Boolean appOnly, Boolean useCachedToken, Boolean reloadConfig)
              
ResultType  : Error
Identity    : Microsoft.Exchange.Security.OAuth.ValidationResultNodeId
IsValid     : True
ObjectState : New