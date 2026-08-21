param(
    [string]$OutputDirectory = "",
    [string]$NodeRuntimePath = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repoRoot "dist\自动下料程序_独立版"
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)

if (Test-Path -LiteralPath $OutputDirectory) {
    throw "输出目录已存在，请指定一个新的 -OutputDirectory：$OutputDirectory"
}

$engineSource = Join-Path $repoRoot "src\build_cutting_plan.mjs"
$launcherSource = Join-Path $repoRoot "src\AutoCuttingLauncher.cs"
$exampleSource = Join-Path $repoRoot "examples"
$usageSource = Join-Path $repoRoot "docs\独立版使用说明.md"
$licenseSource = Join-Path $repoRoot "LICENSE"
$packageSource = Join-Path $repoRoot "package.json"
$lockSource = Join-Path $repoRoot "package-lock.json"

foreach ($requiredPath in @($engineSource, $launcherSource, $exampleSource, $usageSource, $licenseSource, $packageSource, $lockSource)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "缺少构建所需文件：$requiredPath"
    }
}

if ([string]::IsNullOrWhiteSpace($NodeRuntimePath)) {
    $nodeCommand = Get-Command node.exe -ErrorAction SilentlyContinue
    if (-not $nodeCommand) {
        throw "找不到 node.exe。请安装 Node.js，或通过 -NodeRuntimePath 指定运行时路径。"
    }
    $NodeRuntimePath = $nodeCommand.Source
}
$NodeRuntimePath = (Resolve-Path -LiteralPath $NodeRuntimePath).Path

$npmCommand = Get-Command npm.cmd -ErrorAction SilentlyContinue
if (-not $npmCommand) {
    throw "找不到 npm.cmd。请安装 Node.js。"
}

$cscCandidates = @(
    (Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"),
    (Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe")
)
$cscPath = $cscCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $cscPath) {
    throw "找不到 .NET Framework csc.exe。"
}

$engineDirectory = Join-Path $OutputDirectory "engine"
$runtimeDirectory = Join-Path $OutputDirectory "runtime"
$examplesDirectory = Join-Path $OutputDirectory "examples"
New-Item -ItemType Directory -Force -Path $engineDirectory, $runtimeDirectory, $examplesDirectory | Out-Null

Copy-Item -LiteralPath $engineSource -Destination (Join-Path $engineDirectory "build_cutting_plan.mjs")
Copy-Item -LiteralPath $NodeRuntimePath -Destination (Join-Path $runtimeDirectory "node.exe")
Copy-Item -LiteralPath $usageSource -Destination (Join-Path $OutputDirectory "使用说明.md")
Copy-Item -LiteralPath $licenseSource -Destination (Join-Path $OutputDirectory "LICENSE")
Copy-Item -LiteralPath $packageSource -Destination (Join-Path $OutputDirectory "package.json")
Copy-Item -LiteralPath $lockSource -Destination (Join-Path $OutputDirectory "package-lock.json")
Copy-Item -LiteralPath (Join-Path $exampleSource "下料工艺卡_示例.xlsx") -Destination $examplesDirectory

Push-Location $OutputDirectory
try {
    & $npmCommand.Source install --omit=dev --ignore-scripts --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) {
        throw "公开依赖安装失败，退出码：$LASTEXITCODE"
    }
}
finally {
    Pop-Location
}

$launcherOutput = Join-Path $OutputDirectory "自动下料程序.exe"
& $cscPath /nologo /target:winexe /out:$launcherOutput /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll $launcherSource
if ($LASTEXITCODE -ne 0) {
    throw "独立版启动器编译失败，退出码：$LASTEXITCODE"
}

$textExtensions = @(".mjs", ".json", ".md", ".txt", ".cs", ".cmd", ".ps1")
$privatePattern = "@oai|artifact-tool|codex-runtimes|xwechat|wxid_|[A-Za-z]:[\\/]Users[\\/]"
$scanFiles = @(
    Get-ChildItem -LiteralPath $engineDirectory -Recurse -File |
        Where-Object { $textExtensions -contains $_.Extension.ToLowerInvariant() }
    Get-Item -LiteralPath (Join-Path $OutputDirectory "使用说明.md")
    Get-Item -LiteralPath (Join-Path $OutputDirectory "LICENSE")
    Get-Item -LiteralPath (Join-Path $OutputDirectory "package.json")
    Get-Item -LiteralPath (Join-Path $OutputDirectory "package-lock.json")
)
$privateMatches = $scanFiles | Select-String -Pattern $privatePattern -AllMatches
if ($privateMatches) {
    $privateMatches | ForEach-Object { Write-Error $_.ToString() }
    throw "独立版中发现疑似私有依赖或本机路径，已停止打包。"
}

if (Test-Path -LiteralPath (Join-Path $OutputDirectory "node_modules\@oai")) {
    throw "独立版中发现私有依赖目录，已停止打包。"
}

$nodeVersion = (& (Join-Path $runtimeDirectory "node.exe") --version).Trim()
Write-Output "独立版已生成：$OutputDirectory"
Write-Output "Node.js 运行时：$nodeVersion"
Write-Output "启动器：$launcherOutput"
