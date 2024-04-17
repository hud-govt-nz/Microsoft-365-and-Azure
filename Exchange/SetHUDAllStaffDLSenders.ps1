Connect-ExchangeOnline -ShowBanner $false 
$Users =@(
"Culture@hud.govt.nz",
"HUD.Insights@hud.govt.nz",
"HRAssist@hud.govt.nz",
"TeTautiaki@hud.govt.nz",
"HUDInternalComms@hud.govt.nz",
"rainbow@hud.govt.nz",
"Hud.Invoices@hud.govt.nz",
"wahinetoa@hud.govt.nz",
"greengroup@hud.govt.nz",
"coronavirusinfo@hud.govt.nz",
"Facilities@hud.govt.nz",
"HUD_Social_Club@hud.govt.nz",
"peoplesnetwork@hud.govt.nz",
"WhakatipuMauriora@hud.govt.nz", #Mental Health Network
"digitalsupport@hud.govt.nz",
"Annaliese.Riefler@hud.govt.nz",
"Brad.Ward@hud.govt.nz",
"Janine.Crous@hud.govt.nz",
"Kararaina.Calcott-Cribb@hud.govt.nz",
"Gaile.Walker@hud.govt.nz",
"Anne.Shaw@hud.govt.nz",
"Emily.Scarlett@hud.govt.nz",
"Lucy.Hooper@hud.govt.nz",
"Justin.Dahm@hud.govt.nz",
"Deborah.Frost@hud.govt.nz",
"Ben.Dalton@hud.govt.nz",
"Pip.Fox@hud.govt.nz",
"Devon.Heaphy@hud.govt.nz",
"Denise.Sheehan@hud.govt.nz",
"Andrew.Crisp@hud.govt.nz",
"Jo.Hogg@hud.govt.nz",
"Sandra.Mansor@hud.govt.nz",
"Eleisha.Hawkins@hud.govt.nz"
)

Set-DynamicDistributionGroup -Identity "Test Dynamic DL" -AcceptMessagesOnlyFrom $Users