Write-Host
"Enter the hostname of the Exchange 2007 server:
(Example: Exch07a)"

$host = Read-Host

Write-Host 
"Enter the legacy FQDN of the Exchange 2007 server:
(Example: Exch07a.contoso.com)"

$fqdn = Read-Host 

Write-Host 
"********************
This will update the OWA, OAB, UM, ActiveSync, Web Services, and Outlook Anywhere Virtual Directories.
If you see any errors, read them and take appropriate action.
To escape and make no changes, Press CTRL+C

You have selected $fqdn as the legacy namespace.
********************"

Set-OwaVirtualDirectory -Identity "$host\OWA (Default Web Site)" -ExternalUrl https://$fqdn/owa

Set-OabVirtualDirectory -Identity "$host\OAB (Default Web Site)" -InternalUrl https://$fqdn/oab -ExternalUrl https://$fqdn/oab

Set-ActiveSyncVirtualDirectory –Identity “$host\Microsoft-Server-ActiveSync (Default Web Site)” –ExternalUrl $Null –InternalUrl https://$fqdn/Microsoft-Server-ActiveSync

Set-WebServicesVirtualDirectory –Identity “$host\EWS (Default Web Site)” -InternalUrl https://$fqdn/ews/exchange.asmx –ExternalUrl https://$fqdn/EWS/Exchange.asmx

Set-UMVirtualDirectory -Identity "UnifiedMessaging (Default Web Site)" –InternalUrl https://$fqdn/UnifiedMessaging/services.asmx –ExternalUrl https://$fqdn/UnifiedMessaging/services.asmx

Set-OutlookAnywhere -Identity "$host\Rpc (Default WebSite)" -IISAuthenticationMethods Basic,Ntlm