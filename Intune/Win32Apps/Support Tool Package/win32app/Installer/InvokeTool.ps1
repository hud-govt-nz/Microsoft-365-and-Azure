$ScriptFromGitHub = Invoke-WebRequest https://raw.githubusercontent.com/AshForde/HUD-Support-Tool/main/DigitalSupportCommonTasks.ps1
Invoke-Expression $($ScriptFromGitHub.Content)