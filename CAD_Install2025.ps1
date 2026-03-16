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

# 新增：安装前清理Windows Update重启项（通过临时BAT文件管理员运行）
Write-Host "`n🧹 正在清理Windows Update重启项..." -ForegroundColor Cyan
# 1. 定义临时BAT文件路径
$tempBatPath = Join-Path $env:TEMP "CleanWU_Reboot.bat"  # 修复BAT文件名空格问题
# 2. 写入BAT文件内容
@"
@echo off
:: 重启修复 - 清理Windows Update强制重启项
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" /f 2>nul
:: 运行完成
"@ | Out-File -FilePath $tempBatPath -Encoding ASCII
# 3. 管理员身份运行BAT文件
Start-Process -FilePath "cmd.exe" -ArgumentList "/c ""$tempBatPath""" -Verb RunAs -Wait
# 4. 删除临时BAT文件
if (Test-Path $tempBatPath) {
    Remove-Item $tempBatPath -Force
}
Write-Host "✅ Windows Update重启项清理完成" -ForegroundColor Green

# 4. 双盘空间校验：C盘≥4GB，D盘≥10GB
function Check-DiskSpace {
    $cDrive = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.Name -eq 'C:\' -and $_.DriveType -eq [System.IO.DriveType]::Fixed }
    $dDrive = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.Name -eq 'D:\' -and $_.DriveType -eq [System.IO.DriveType]::Fixed }

    if (-not $cDrive) { throw "未检测到C盘" }
    if (-not $dDrive) { throw "未检测到D盘" }

    $cFreeGB = [math]::Round($cDrive.AvailableFreeSpace / 1GB, 1)
    $dFreeGB = [math]::Round($dDrive.AvailableFreeSpace / 1GB, 1)

    if ($cFreeGB -lt 4) { throw "C盘空间不足4GB（当前可用：$cFreeGB GB）" }
    if ($dFreeGB -lt 10) { throw "D盘空间不足10GB（当前可用：$dFreeGB GB）" }

    return @{ C = $cFreeGB; D = $dFreeGB }
}

# 5. 核心配置（适配2025版）
$DownloadUrl  = "http://115.191.18.103:5244/d/%E7%A7%BB%E5%8A%A8/CAD_Shell/AutoCAD_2025_Shell_YJ.zip"
$ZipPath      = Join-Path $env:TEMP "AutoCAD_2025_Shell_YJ.zip"
$FinalDir     = "D:\AutoCAD_2025_Shell_YJ"
$SetupBatPath = Join-Path $FinalDir "Install AutoCAD 2025.bat"
$ImageDir     = Join-Path $FinalDir "image"
$logPath      = Join-Path $FinalDir "install_log.txt"
$SourceAcadExe = Join-Path $FinalDir "acad.exe"
$TargetAcadDir = "D:\Autodesk\AutoCAD 2025"
$TargetAcadExe = Join-Path $TargetAcadDir "acad.exe"

# 确保目标目录存在
if (-not (Test-Path $FinalDir)) {
    New-Item -Path $FinalDir -ItemType Directory -Force | Out-Null
}

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
            try {
                $webClient.DownloadFile($url, $path)
            }
            catch {
                Write-Error "下载失败：$($_.Exception.Message)"
            }
            finally {
                $webClient.Dispose()
            }
        } -ArgumentList $DownloadUrl, $ZipPath

        $startTime = Get-Date
        $lastSize = 0
        $totalSizeEst = 3.12 * 1GB

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

        # 检查下载任务是否出错
        $jobResult = Receive-Job $job -Wait
        if ($jobResult -match "下载失败") {
            Write-Host "`n❌ $jobResult" -ForegroundColor Red
            Read-Host "按任意键退出"
            exit 1
        }
        Remove-Job $job -Force

        if (-not (Test-Path $ZipPath) -or (Get-Item $ZipPath).Length -eq 0) {
            Write-Host "`n❌ 下载的安装包为空或损坏" -ForegroundColor Red
            Read-Host "按任意键退出"
            exit 1
        }

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

    # 解压到D盘（修复版：跳过压缩包内顶层文件夹，直接提取内容到FinalDir）
    Write-Host "`n📦 正在解压中耐心等待..." -ForegroundColor Yellow
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zipFile = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $entries = $zipFile.Entries
        $total = $entries.Count
        $current = 0

        # 找到压缩包内的顶层文件夹名（如 AutoCAD_2025_Shell_YJ\）
        $topFolder = $entries | Where-Object { $_.FullName -match '^[^/]+/$' } | Select-Object -First 1 -ExpandProperty FullName
        if (-not $topFolder) { $topFolder = "" }

        foreach ($entry in $entries) {
            if (-not [string]::IsNullOrEmpty($entry.Name)) {
                # 核心修复：如果路径以顶层文件夹开头，就截掉这部分
                $relativePath = if ($entry.FullName.StartsWith($topFolder)) {
                    $entry.FullName.Substring($topFolder.Length)
                } else {
                    $entry.FullName
                }
                $targetPath = Join-Path $FinalDir $relativePath
                $targetDir = Split-Path $targetPath -Parent
                
                if (-not (Test-Path $targetDir)) {
                    [System.IO.Directory]::CreateDirectory($targetDir) | Out-Null
                }
                # 仅解压文件（跳过目录条目）
                if (-not $entry.FullName.EndsWith("/")) {
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)
                }
            }
            $current++
            $percent = [math]::Round(($current / $total) * 100, 1)
            Write-Host "`r解压进度：$percent% ($current/$total)" -NoNewline
        }
        $zipFile.Dispose()
        Write-Host "`n✅ 解压完成！" -ForegroundColor Green
        
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
    Write-Host "❌ 未找到安装文件 Install AutoCAD 2025.bat" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}
Write-Host "✅ 所有验证通过，开始安装" -ForegroundColor Green

# 8. 安装
try {
    Write-Host "🔄 正在启动 AutoCAD 2025 安装，请耐心等待安装完成..." -ForegroundColor Cyan
    $installProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c start """" ""$SetupBatPath"" >> `"$logPath`" 2>&1" -Verb RunAs -PassThru

    # 第一步：等待Installer.exe进程完全结束（静默等待，无提示）
    $installerWaitCount = 0
    $maxInstallerWait = 360  # 最多等待30分钟（60*5秒）
    do {
        Start-Sleep -Seconds 5
        $installerProcess = Get-Process -Name "Installer" -ErrorAction SilentlyContinue
        $installerWaitCount++
        
        # 超时保护：防止进程一直存在导致卡死
        if ($installerWaitCount -ge $maxInstallerWait) {
            Write-Host "`n⚠️ Installer进程等待超时，强制继续检测" -ForegroundColor Yellow
            break
        }
    } while ($installerProcess -ne $null)

    # 进程结束后显示检测提示，仅检测目标目录是否存在acad.exe
    Write-Host "🔍 安装检测中..." -ForegroundColor Cyan
    $cad2025Installed = $false
    $fileWaitCount = 0
    $maxFileWait = 240  # 最多等待20分钟（48*5秒）

    # 第二步：仅检测$TargetAcadDir中是否存在acad.exe
    do {
        Start-Sleep -Seconds 5
        $fileWaitCount++
        
        # 核心检测：目标目录是否存在acad.exe文件
        if (Test-Path $TargetAcadExe -PathType Leaf) {
            $cad2025Installed = $true
        }

        # 超时保护：防止文件一直不存在导致卡死
        if ($fileWaitCount -ge $maxFileWait) {
            Write-Host "`n❌ 检测超时：$TargetAcadExe 文件未找到" -ForegroundColor Red
            $userChoice = Read-Host "是否强制继续激活流程？(Y/N)"
            if ($userChoice -eq "Y" -or $userChoice -eq "y") {
                $cad2025Installed = $true
            } else {
                exit 1
            }
        }
    } while (-not $cad2025Installed)

    Write-Host "✅ 安装检测完成" -ForegroundColor Green
    Start-Sleep -Seconds 3

    Write-Host "`n🎉 安装完成！" -ForegroundColor Green
    Write-Host "ℹ️ 日志已保存到：$logPath" -ForegroundColor Cyan

    if (Test-Path $SetupBatPath -PathType Leaf) {
        Remove-Item $SetupBatPath -Force
        Write-Host "🗑️ 已自动打扫安装批处理文件" -ForegroundColor Green
    }

    Write-Host "`n📤 开始激活中..." -ForegroundColor Cyan
    if (Test-Path $SourceAcadExe -PathType Leaf) {
        # 确保目标目录存在
        if (-not (Test-Path $TargetAcadDir -PathType Container)) {
            New-Item -Path $TargetAcadDir -ItemType Directory -Force | Out-Null
            Write-Host "📁 已创建目标目录：$TargetAcadDir" -ForegroundColor Yellow
        }
        
        # 备份原文件（可选）
        if (Test-Path $TargetAcadExe -PathType Leaf) {
            Copy-Item -Path $TargetAcadExe -Destination "$TargetAcadExe.bak" -Force -ErrorAction SilentlyContinue
        }
        # 执行替换
        Copy-Item -Path $SourceAcadExe -Destination $TargetAcadExe -Force
        Write-Host "✅ 已激活完成" -ForegroundColor Green

        Remove-Item -Path $SourceAcadExe -Force
        Write-Host "🗑️ 已打扫源文件" -ForegroundColor Green
    }
    else {
        Write-Host "⚠️ 未找到源文件：$SourceAcadExe，跳过复制替换" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "`n❌ 安装/文件替换失败：$($_.Exception.Message)" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}

# 9. 脚本自删除
$scriptPath = $MyInvocation.MyCommand.Definition
if (Test-Path $scriptPath -PathType Leaf) {
    try {
        Remove-Item $scriptPath -Force
        Write-Host "🗑️ 已自动打扫脚本文件" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️ 脚本自删除失败：$($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Read-Host "`n所有操作结束，按任意键退出"
