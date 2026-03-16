# 1. 兼容编码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# 2. 兼容TLS
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
catch {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls
}

# 3. 强制管理员运行
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process powershell.exe "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 4. 双盘空间校验：C盘≥2GB，D盘≥10GB
function Check-DiskSpace {
    $cDrive = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.Name -eq 'C:\' -and $_.DriveType -eq [System.IO.DriveType]::Fixed }
    $dDrive = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.Name -eq 'D:\' -and $_.DriveType -eq [System.IO.DriveType]::Fixed }

    if (-not $cDrive) { throw "未检测到C盘" }
    if (-not $dDrive) { throw "未检测到D盘" }

    $cFreeGB = [math]::Round($cDrive.AvailableFreeSpace / 1GB, 1)
    $dFreeGB = [math]::Round($dDrive.AvailableFreeSpace / 1GB, 1)

    if ($cFreeGB -lt 4) { throw "C盘空间不足4GB（需存放压缩包）" }
    if ($dFreeGB -lt 10) { throw "D盘空间不足10GB（安装程序）" }

    return @{ C = $cFreeGB; D = $dFreeGB }
}

# 5. 核心配置（适配2024版：Img→image、666.lnk→Install AutoCAD 2024_1.bat）
$DownloadUrl  = "http://115.191.18.103:5244/d/%E7%A7%BB%E5%8A%A8/CAD_Shell/AutoCAD_2024_Shell_YJ.zip"
$ZipPath      = Join-Path $env:TEMP "AutoCAD_2024_Shell_YJ.zip"  # C盘%temp%下载
$FinalDir     = "D:\AutoCAD_2024_Shell_YJ"                      # D盘固定解压目录
$SetupBatPath = Join-Path $FinalDir "Install AutoCAD 2024_1.bat" # 替换666.lnk为指定bat文件
$ImageDir     = Join-Path $FinalDir "image"                      # Img→image
$logPath      = Join-Path $FinalDir "install_log.txt"            # 日志保存到解压文件夹
$SourceAcadExe = Join-Path $FinalDir "acad.exe"                  # 源acad.exe路径
$TargetAcadDir = "D:\Autodesk\AutoCAD 2024"                      # 目标替换目录
$TargetAcadExe = Join-Path $TargetAcadDir "acad.exe"             # 目标acad.exe路径

# 6. 优先检测D盘是否已有完整安装文件
Write-Host "`n🔍 磁盘检测中..." -ForegroundColor Cyan
$batExists = Test-Path $SetupBatPath -PathType Leaf
$imageDirExists = Test-Path $ImageDir -PathType Container
$skipAll = $false

if ($batExists -and $imageDirExists) {
    $imageItems = Get-ChildItem -Path $ImageDir -Recurse -ErrorAction SilentlyContinue
    if ($imageItems -and $imageItems.Count -gt 0) {
        Write-Host "✅ 检测到完整安装文件，跳过下载+解压" -ForegroundColor Green
        $skipAll = $true
    }
}

if (-not $skipAll) {
    # 双盘空间校验
    try {
        $diskInfo = Check-DiskSpace
        Write-Host "`n🔍 磁盘剩余空间：C盘 $($diskInfo.C) GB | D盘 $($diskInfo.D) GB" -ForegroundColor Cyan
    }
    catch {
        Write-Host "❌ $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "按任意键退出"
        exit 1
    }

    Write-Host "`n❌ 未检测到完整文件，开始下载" -ForegroundColor Cyan
    
    if (-not (Test-Path $ZipPath -PathType Leaf)) {
        Write-Host "`n📥 正在下载安装包（耐心等待）..." -ForegroundColor Yellow
        
        $job = Start-Job -ScriptBlock {
            param($url, $path)
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $webClient = New-Object System.Net.WebClient
            $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
            $webClient.DownloadFile($url, $path)
            $webClient.Dispose()
        } -ArgumentList $DownloadUrl, $ZipPath

        $startTime = Get-Date
        $lastSize = 0
        $totalSizeEst = 2.94 * 1GB

        while ($job.State -eq "Running") {
            if (Test-Path $ZipPath -PathType Leaf) {
                $currentSize = (Get-Item $ZipPath).Length
                $currentTime = Get-Date
                $elapsed = ($currentTime - $startTime).TotalSeconds
                $avgSpeed = if ($elapsed -gt 0) { [math]::Round(($currentSize / 1024) / $elapsed, 1) } else { 0 }
                $downloadedGB = [math]::Round($currentSize / 1GB, 2)

                $remainingSize = $totalSizeEst - $currentSize
                $remainingSec = if ($avgSpeed -gt 0) { [math]::Round(($remainingSize / 1024) / $avgSpeed, 1) } else { 0 }
                $remainingMin = [math]::Round($remainingSec / 60, 1)
                $remainingMin = if ($remainingMin -lt 0.1) { 0.1 } else { $remainingMin }

                Write-Host "`r📦 已下载：$downloadedGB GB | 速度：$avgSpeed KB/s | 剩余时间：$remainingMin 分钟" -NoNewline -ForegroundColor Yellow
                $lastSize = $currentSize
            }
            Start-Sleep -Milliseconds 500
        }

        Receive-Job $job -Wait | Out-Null
        Remove-Job $job -Force

        $endTime = Get-Date
        $totalTime = ($endTime - $startTime).TotalSeconds
        $finalSize = (Get-Item $ZipPath).Length
        $avgSpeed = [math]::Round($finalSize / 1024 / $totalTime, 1)
        $finalGB = [math]::Round($finalSize / 1GB, 2)
        Write-Host "`n✅ 下载完成！| 总大小：$finalGB GB | 平均速度：$avgSpeed KB/s" -ForegroundColor Green
    }
    else {
        Write-Host "✅ C盘已存在压缩包，跳过下载" -ForegroundColor Green
    }

    # 解压到D盘（仅修改此段，适配图片中的解压结构）
    Write-Host "`n📦 正在解压中耐心等待..." -ForegroundColor Yellow
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zipFile = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $entries = $zipFile.Entries
        $total = $entries.Count
        $current = 0

        foreach ($entry in $entries) {
            if (-not [string]::IsNullOrEmpty($entry.Name)) {
                # 核心修改：直接解压到FinalDir根目录，不做任何路径分割
                $targetPath = Join-Path $FinalDir $entry.FullName
                $targetDir = Split-Path $targetPath -Parent
                if (-not (Test-Path $targetDir)) {
                    [System.IO.Directory]::CreateDirectory($targetDir) | Out-Null
                }
                # 仅解压文件（跳过目录条目，避免重复创建目录）
                if (-not $entry.FullName.EndsWith("/")) {
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)
                }
            }
            $current++
            $percent = [math]::Round(($current / $total) * 100, 1)
            Write-Host "`r解压进度：$percent% ($current/$total)" -NoNewline
        }
        $zipFile.Dispose()  # 释放压缩包句柄
        Write-Host "`n✅ 解压完成！" -ForegroundColor Green
        
        # 自动删除C盘压缩包
        if (Test-Path $ZipPath -PathType Leaf) {
            Remove-Item $ZipPath -Force
            Write-Host "🗑️ 已自动打扫压缩包" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "`n❌ 解压失败：$($_.Exception.Message)" -ForegroundColor Red
        Read-Host "按任意键退出"
        exit 1
    }
}

# 7. 安装前验证
Write-Host "`n🔍 安装前二次验证..." -ForegroundColor Cyan
if (-not (Test-Path $ImageDir -PathType Container)) {
    Write-Host "❌ image文件夹不存在" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}
$imageItems = Get-ChildItem -Path $ImageDir -Recurse -ErrorAction SilentlyContinue
if ($imageItems -eq $null -or $imageItems.Count -eq 0) {
    Write-Host "❌ image文件夹为空" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}
if (-not (Test-Path $SetupBatPath)) {
    Write-Host "❌ 未找到安装文件 Install AutoCAD 2024_1.bat" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}
Write-Host "✅ 所有验证通过，开始安装" -ForegroundColor Green

# 8. 安装
try {
    Write-Host "🔄 正在启动 AutoCAD 2024 安装，请耐心等待安装完成..." -ForegroundColor Cyan
    $installProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c start """" ""$SetupBatPath"" >> `"$logPath`" 2>&1" -Verb RunAs -PassThru

    # 循环等待 Installer.exe 进程结束（真正等待安装完成）
    do {
        Start-Sleep -Seconds 5
        $installerProcess = Get-Process -Name "Installer" -ErrorAction SilentlyContinue
    } while ($installerProcess -ne $null)

    # 再等3秒，确保安装后清理动作完成
    Start-Sleep -Seconds 3

    Write-Host "`n🎉 安装完成！" -ForegroundColor Green
    Write-Host "ℹ️ 日志已保存到：$logPath" -ForegroundColor Cyan

    # 安装完成后自动删除 Install AutoCAD 2024_1.bat
    if (Test-Path $SetupBatPath -PathType Leaf) {
        Remove-Item $SetupBatPath -Force
        Write-Host "🗑️ 已自动打扫安装批处理文件" -ForegroundColor Green
    }

    # 新增：复制acad.exe到目标目录并替换，然后删除源文件
    Write-Host "`n📤 开始激活中..." -ForegroundColor Cyan
    if (Test-Path $SourceAcadExe -PathType Leaf) {
        # 检查目标目录是否存在，不存在则创建
        if (-not (Test-Path $TargetAcadDir -PathType Container)) {
            New-Item -Path $TargetAcadDir -ItemType Directory -Force | Out-Null
            Write-Host "📁 已创建目标目录：$TargetAcadDir" -ForegroundColor Yellow
        }
        
        # 复制并覆盖（强制替换）
        Copy-Item -Path $SourceAcadExe -Destination $TargetAcadExe -Force
        Write-Host "✅ 已激活完成" -ForegroundColor Green

        # 删除源文件
        Remove-Item -Path $SourceAcadExe -Force
        Write-Host "🗑️ 已打扫源文件" -ForegroundColor Green
    }
    else {
        Write-Host "⚠️ 未找到源文件：$SourceAcadExe，跳过复制替换" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "`n❌ 安装/文件替换失败：$($_.Exception.Message)" -ForegroundColor Red
}

# 9. 脚本自删除
$scriptPath = $MyInvocation.MyCommand.Definition
if (Test-Path $scriptPath -PathType Leaf) {
    Remove-Item $scriptPath -Force
    Write-Host "🗑️ 已自动打扫脚本文件" -ForegroundColor Green
}

Read-Host "`n所有操作结束，按任意键退出"
