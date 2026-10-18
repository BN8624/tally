# Tally를 빨리 시작하는 폴더형 실행 파일로 빌드하고 배포용 zip을 만듭니다.
# 사용법: powershell -ExecutionPolicy Bypass -File .\build.ps1
Set-Location $PSScriptRoot

$python = "build\venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    python -m venv build\venv
    if ($LASTEXITCODE -ne 0) { throw "빌드용 가상환경을 만들지 못했습니다. Python 3.11 이상이 설치되어 있는지 확인하세요." }
}
& $python -m pip install --disable-pip-version-check -q pandas==3.0.3 numpy==2.5.1 openpyxl==3.1.5 pyinstaller==6.22.3
if ($LASTEXITCODE -ne 0) { throw "빌드에 필요한 패키지를 설치하지 못했습니다." }

# 한 파일(onefile)은 실행할 때마다 압축을 풀어 느리므로 폴더형(onedir)으로 만듭니다.
# 앱이 쓰지 않는 인터넷 암호화(ssl·hashlib), sqlite, numpy 포트란 도구는 빼서 용량을 줄입니다.
$arguments = @(
    "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir", "--windowed", "--name", "Tally",
    "--workpath", "build\pyinstaller", "--specpath", "build", "--distpath", "dist"
)
foreach ($module in "_ssl", "ssl", "_hashlib", "sqlite3", "_sqlite3", "numpy.f2py") {
    $arguments += @("--exclude-module", $module)
}
& $python @arguments app.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller 빌드에 실패했습니다." }

$app = Join-Path $PSScriptRoot "dist\Tally"
# Tcl 시간대 데이터는 Tk 화면에서 쓰지 않으므로 지웁니다.
Remove-Item (Join-Path $app "_internal\_tcl_data\tzdata") -Recurse -Force -ErrorAction SilentlyContinue
if (Test-Path (Join-Path $app "data")) {
    throw "dist\Tally\data에 업체 설정이 있습니다. 실제 업체명이 배포되지 않도록 지운 뒤 다시 실행하세요."
}

$commit = (git rev-parse --short HEAD).Trim()
if (git status --porcelain) { $commit += "-dirty" }
$zip = Join-Path $PSScriptRoot ("dist\Tally_{0}_{1}.zip" -f (Get-Date -Format yyyyMMdd), $commit)
if (Test-Path $zip) { Remove-Item $zip -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::CreateFromDirectory($app, $zip, [IO.Compression.CompressionLevel]::Optimal, $true)

$folderMB = (Get-ChildItem $app -Recurse -File | Measure-Object Length -Sum).Sum / 1MB
"완료: 폴더 {0:N1}MB, zip {1:N1}MB, {2}" -f $folderMB, ((Get-Item $zip).Length / 1MB), $zip
