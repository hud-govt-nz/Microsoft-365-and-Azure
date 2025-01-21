#https://www.experts-exchange.com/articles/33601/How-to-Stop-emails-to-onmicrosft-com-O365-domains-from-External-World.html
#https://docs.microsoft.com/en-us/powershell/exchange/disable-access-to-exchange-online-powershell?view=exchange-ps
#https://docs.microsoft.com/en-gb/exchange/mail-flow-best-practices/manage-mail-flow-using-third-party-cloud


#List all inbound connectors configured by HCW
$onpremorg = Get-OnPremisesOrganization | Select OrganizationGuid, InboundConnector | Where {$_.InboundConnector -ne $null}
$onpremorg

#New inbound connetor
New-InboundConnector `
    -Name “Block Direct Delivery to *@tenant.OnMicrosoft.com Alias” `
    -ConnectorType Partner `
    -SenderDomains * `
    -TlsSenderCertificateName (Get-InboundConnector $onpremorg[0].InboundConnector).TlsSenderCertificateName `
    -RestrictDomainsToCertificate $True `
    -RequireTls $True
