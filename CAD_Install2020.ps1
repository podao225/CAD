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
$tempBatPath = Join-Path $env:TEMP "CleanWU Reboot.bat"
# 2. 写入BAT文件内容
@"
@echo off
:: 重启修复 - 清理Windows Update强制重启项
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" /f
:: 运行完成
"@ | Out-File -FilePath $tempBatPath -Encoding ASCII
# 3. 管理员身份运行BAT文件
Start-Process -FilePath "cmd.exe" -ArgumentList "/c ""$tempBatPath""" -Verb RunAs -Wait
# 4. 删除临时BAT文件
if (Test-Path $tempBatPath) {
    Remove-Item $tempBatPath -Force
}
Write-Host "✅ Windows Update重启项清理完成" -ForegroundColor Green

# 4. 双盘空间校验：C盘≥2GB，D盘≥10GB
function Check-DiskSpace {
    $cDrive = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.Name -eq 'C:\' -and $_.DriveType -eq [System.IO.DriveType]::Fixed }
    $dDrive = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.Name -eq 'D:\' -and $_.DriveType -eq [System.IO.DriveType]::Fixed }

    if (-not $cDrive) { throw "未检测到C盘" }
    if (-not $dDrive) { throw "未检测到D盘" }

    $cFreeGB = [math]::Round($cDrive.AvailableFreeSpace / 1GB, 1)
    $dFreeGB = [math]::Round($dDrive.AvailableFreeSpace / 1GB, 1)

    if ($cFreeGB -lt 2) { throw "C盘空间不足2GB（需存放压缩包）" }
    if ($dFreeGB -lt 10) { throw "D盘空间不足10GB（安装程序）" }

    return @{ C = $cFreeGB; D = $dFreeGB }
}

# 5. 核心配置（固定D盘为解压/安装目录，日志保存到解压文件夹）
$DownloadUrl  = "http://115.191.18.103:5244/d/%E7%A7%BB%E5%8A%A8/CAD_Shell/AutoCAD_2020_Shell_YJ.zip"
$ZipPath      = Join-Path $env:TEMP "AutoCAD_2020_Shell_YJ.zip"  # C盘%temp%下载
$FinalDir     = "D:\AutoCAD_2020_Shell_YJ"                      # D盘固定解压目录
$SetupLnkPath = Join-Path $FinalDir "666.lnk"
$ImgDir       = Join-Path $FinalDir "Img"
$logPath      = Join-Path $FinalDir "install_log.txt"           # 日志保存到解压文件夹

# 6. 优先检测D盘是否已有完整安装文件
Write-Host "`n🔍 磁盘检测中..." -ForegroundColor Cyan
$lnkExists = Test-Path $SetupLnkPath -PathType Leaf
$imgDirExists = Test-Path $ImgDir -PathType Container
$skipAll = $false

if ($lnkExists -and $imgDirExists) {
    $imgItems = Get-ChildItem -Path $ImgDir -Recurse -ErrorAction SilentlyContinue
    if ($imgItems -and $imgItems.Count -gt 0) {
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
        $totalSizeEst = 1.88 * 1GB

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

    # 解压到D盘
    Write-Host "`n📦 正在解压中耐心等待..." -ForegroundColor Yellow
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zipFile = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        $entries = $zipFile.Entries
        $total = $entries.Count
        $current = 0

        foreach ($entry in $entries) {
            if (-not [string]::IsNullOrEmpty($entry.Name)) {
                $entryParts = $entry.FullName -split '[\\/]'
                if ($entryParts.Count -gt 1) {
                    $innerPath = $entryParts[1..($entryParts.Count-1)] -join '\'
                    $targetPath = Join-Path $FinalDir $innerPath
                } else {
                    $targetPath = Join-Path $FinalDir $entry.FullName
                }
                $targetDir = Split-Path $targetPath -Parent
                if (-not (Test-Path $targetDir)) {
                    [System.IO.Directory]::CreateDirectory($targetDir) | Out-Null
                }
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)
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
if (-not (Test-Path $ImgDir -PathType Container)) {
    Write-Host "❌ Img文件夹不存在" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}
$imgItems = Get-ChildItem -Path $ImgDir -Recurse -ErrorAction SilentlyContinue
if ($imgItems -eq $null -or $imgItems.Count -eq 0) {
    Write-Host "❌ Img文件夹为空" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}
if (-not (Test-Path $SetupLnkPath)) {
    Write-Host "❌ 未找到安装文件" -ForegroundColor Red
    Read-Host "按任意键退出"
    exit 1
}
Write-Host "✅ 所有验证通过，开始安装" -ForegroundColor Green

# 8. 安装
try {
    # 1. 启动安装程序
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c start """" ""$SetupLnkPath"" >> `"$logPath`" 2>&1" -Verb RunAs -Wait
    Write-Host "`n🎉 安装程序启动完成！" -ForegroundColor Green
    Write-Host "ℹ️ 安装日志已保存到：$logPath" -ForegroundColor Cyan

    # 2. 检测AutoCAD主程序是否生成
    $cadExePath = "D:\Autodesk\AutoCAD 2020\acad.exe"
    $isCadExeExists = Test-Path $cadExePath -PathType Leaf

    if ($isCadExeExists) {
        # 2.1 主程序生成成功：直接清理安装文件
        Write-Host "✅ 检测到AutoCAD主程序已生成：$cadExePath" -ForegroundColor Green
        
        # 清理666.lnk文件
        if (Test-Path $SetupLnkPath -PathType Leaf) {
            Remove-Item $SetupLnkPath -Force
            Write-Host "🗑️ 已自动删除安装快捷方式 666.lnk" -ForegroundColor Green
        }
    }
    else {
        # 2.2 主程序未生成：弹出人工选择
        Write-Host "❌ 未检测到AutoCAD主程序：$cadExePath" -ForegroundColor Red
        
        # 循环获取有效输入
        do {
            $userChoice = Read-Host "`n是否继续执行后续操作？[1=继续清理文件 / 2=退出并删除脚本]"
            switch ($userChoice) {
                "1" {
                    # 选择继续：仅执行文件清理
                    Write-Host "🔍 选择继续，开始清理安装文件..." -ForegroundColor Cyan
                    if (Test-Path $SetupLnkPath -PathType Leaf) {
                        Remove-Item $SetupLnkPath -Force
                        Write-Host "🗑️ 已自动删除安装快捷方式 666.lnk" -ForegroundColor Green
                    }
                    break
                }
                "2" {
                    # 选择退出：删除脚本并结束程序
                    Write-Host "🛑 选择退出，正在删除脚本文件..." -ForegroundColor Red
                    $scriptPath = $MyInvocation.MyCommand.Definition
                    if (Test-Path $scriptPath -PathType Leaf) {
                        Remove-Item $scriptPath -Force
                        Write-Host "🗑️ 已自动删除脚本文件" -ForegroundColor Green
                    }
                    Write-Host "🔚 程序已结束" -ForegroundColor Red
                    exit 0
                }
                default {
                    Write-Host "❌ 输入无效！请输入 1 或 2" -ForegroundColor Red
                }
            }
        } while ($userChoice -notin @("1", "2"))
    }
}
catch {
    # 3. 安装启动失败的异常处理
    Write-Host "`n❌ 安装程序启动失败：$($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# 9. 脚本自删除
$scriptPath = $MyInvocation.MyCommand.Definition
if (Test-Path $scriptPath -PathType Leaf) {
    Remove-Item $scriptPath -Force
    Write-Host "🗑️ 已自动打扫脚本文件" -ForegroundColor Green
}

Read-Host "`n所有操作结束，按任意键退出"
