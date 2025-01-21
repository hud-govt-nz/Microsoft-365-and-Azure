# https://blog.onevinn.com/using-powershell-to-get-wildcard-certificate-from-lets-encrypt

Install-Module -Name Posh-ACME

New-PACertificate kuhu.hud.govt.nz -AcceptTOS -Contact digitalsupport@hud.govt.nz -DnsPlugin AcmeDns -PluginArgs @{ACMEServer='auth.acme-dns.io'} -Install

Get-PACertificate | Format-List