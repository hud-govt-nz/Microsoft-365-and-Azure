library(AzureStor) 
library(AzureRMR)

resource <- "https://sysintel.blob.core.windows.net"
container <- "sandbox"
token <- get_azure_token("https://storage.azure.com",
                         tenant = "9e9b3020-3d38-48a6-9064-373bc7b156dc",
                         app = "c6c4300b-9ff3-4946-8f30-e0aa59bdeaf5")
endp_key <- storage_endpoint(resource, token = token)
storage_container(endp_key, container)
list_blobs(cont)
install.packages("languageserver")