$cgm = @("sip:norman.niro@hud.govt.nz","sip:janine.crous@hud.govt.nz","sip:aparna.sreekumar@hud.govt.nz","sip:ashley.forde@hud.govt.nz")
Set-CsUserCallingSettings -Identity AppAdmin@mhud.onmicrosoft.com -CallGroupOrder InOrder -CallGroupTargets $cgm
Set-CsUserCallingSettings -Identity AppAdmin@mhud.onmicrosoft.com -IsForwardingEnabled $true -ForwardingType Immediate -ForwardingTargetType Group