$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
python -m venv .venv
if ($LASTEXITCODE -ne 0) { throw 'Não foi possível criar o ambiente Python.' }
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw 'Não foi possível instalar as dependências.' }
& .\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name TESTEB2B --add-data 'ui;ui' --collect-all webview --collect-all msal --collect-all pymsalruntime app.py
if ($LASTEXITCODE -ne 0) { throw 'Não foi possível gerar o executável.' }
Copy-Item -LiteralPath .env.example -Destination dist\.env.example -Force
Write-Host 'Pronto: dist\TESTEB2B.exe. Copie .env.example para .env e preencha a configuração.'
