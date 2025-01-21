$regFilePath = Join-Path -Path $PSScriptRoot -ChildPath "your_file.reg"
Start-Process -FilePath "reg.exe" -ArgumentList "import '$regFilePath'" -NoNewWindow -Wait