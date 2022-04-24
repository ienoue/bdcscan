@setlocal enabledelayedexpansion&set a=%*&(if defined a set a=!a:"=\"!&set a=!a:\\"=\\\"!&set a=!a:'=''!)&powershell/c $i=$input;iex ('$i^|^&{Param([switch]$silent);$PSCommandPath=\"%~f0\";$PSScriptRoot=\"%~dp0";#'+(${%~f0}^|Out-String)+'} '+('!a!'-replace'[$(),;@`{}]','`$0'))&exit/b

$updateFileURL = "http://download.bitdefender.com/updates/update_av32bit/cumulative.zip"
$debugPreference = "Continue"
$elapsedMinutesThreshold = 1440

function Test-Update {
    Param(
        [parameter(Mandatory=$true)]
        [string]$cumulativeFileURL,
        [parameter(Mandatory=$true)]
        [DateTime]$utcUpdateTime
    )
    try{
        [System.Net.HttpWebRequest]$request = [System.Net.WebRequest]::Create($cumulativeFileURL)
        $request.Method ="HEAD"
        $request.Timeout = 10000
        $response = $request.GetResponse()
        $headerUpdateTime = ($response.LastModified).ToUniversalTime()
        $statusCode = [int]$response.StatusCode
        if ($statusCode -eq [int][system.net.httpstatuscode]::ok) {
            if(($headerUpdateTime - $utcUpdateTime).TotalMinutes -ge $elapsedMinutesThreshold) {
                Write-Host "新しいアップデートがありました"
                return $true
            } else {
                Write-Host "定義ファイルは最新です"
                return $false
            }
        }
    } catch [System.Net.WebException] {
        Write-Warning $_.Exception.Message
    } catch {
        Write-Debug ($_.Exception.tostring() + $_.InvocationInfo.PositionMessage)
    } finally {
         if ($response -ne $null) {
            $response.Close()
        }
    }
    Write-Warning "$($cumulativeFileURL)に正常に接続できませんでした"
    return $false
}

function Download-File {
    Param(
        [parameter(Mandatory=$true)]
        [string]$sourceURL,
        [parameter(Mandatory=$true)]
        [string]$destinationPath
    )
        Write-Host "$($sourceURL)のダウンロードを開始します"
    try {
        Import-Module BitsTransfer
        Start-BitsTransfer -Source $sourceURL -Destination $destinationPath -RetryInterval 60 -RetryTimeout 120 -ErrorAction Stop
        Write-Host "ダウンロードが完了しました"
    } catch {
        Write-Warning "ダウンロードに失敗しました"
        throw
    }
}

function Extract-Zip {
    Param(
        [parameter(Mandatory=$true)]
        [string]$zipFilePath,
        [parameter(Mandatory=$true)]
        [string]$bitDefenderPath
    )
    $major = $PSVersionTable.CLRVersion.Major
    $revision = $PSVersionTable.CLRVersion.Revision
    #4.0.30319.17020.NET 4.5 Preview, September 2011
    if (($major -ge 5) -or (($major -ge 4) -and ($revision -ge 17020))) {
        try {
           [void][System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem")
            $zipFile = [System.IO.Compression.ZipFile]::OpenRead($zipFilePath)
            foreach ($entry in $zipFile.Entries) {
                if ($entry.FullName -eq "bdcore.dll") {
                    $isCorePlugin = $true
                    break
                }
            }
            if (!$isCorePlugin) {
                Write-Warning "BitDefenderV10に対応した定義ファイルではありません"
                throw "bdcore.dllが圧縮ファイルに含まれていません"
            }
            $zipFile.Entries | Where-Object{ ($_.FullName.StartsWith("Plugins/", $true, $null)) -or ($_.FullName.EndsWith(".dll", $true, $null)) } | ForEach-Object {
                $destinationPath = Join-Path $bitDefenderPath $_.FullName
                $parentPath = Split-Path $destinationPath -parent
                if (!(Test-Path $parentPath)) {
                    [void](New-Item $parentPath -ItemType Directory)
                }
                if ($destinationPath.EndsWith("\")) {
                    if (!(Test-Path $destinationPath)) {
                        [void](New-Item $destinationPath -ItemType Directory)
                    }
                } else {
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $destinationPath, $true)
                }
            }
        } catch {
            Write-Warning "$($zipFilePath)の展開に失敗しました"
            throw
        } finally {
            if ($zipFile -ne $null) {
                $zipFile.Dispose()
            }
        }
    } else {
        try {
            $shell = New-Object -ComObject Shell.Application
            $zipFile = $shell.NameSpace($zipFilePath)
            foreach ($item in $zipFile.Items()) {
                if ($item.Name -eq "bdcore.dll") {
                    $isCorePlugin = $true
                    break
                }
            }
            if (!$isCorePlugin) {
                Write-Warning "BitDefenderV10に対応した定義ファイルではありません"
                throw "bdcore.dllが圧縮ファイルに含まれていません"
            }
            if (!(Test-Path $bitDefenderPath)) {
                [void](New-Item $bitDefenderPath -ItemType Directory)
            }
            $destinationFolder = $shell.NameSpace($bitDefenderPath)
            $zipFile.Items() | Where-Object{ ($_.IsFolder -and $_.Name -eq "Plugins") -or ($_.Name.EndsWith(".dll", $true, $null)) } | ForEach-Object {
                $destinationFolder.CopyHere($_.Path, 0x14)
            }
        } catch {
            Write-Warning "$($zipFilePath)の解凍に失敗しました"
            throw
        } finally {
            $zipFile, $destinationFolder, $shell | ForEach-Object {
            if ($_ -ne $null) {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($_)
                }
            }
        }
    }
}

function Update-Plugins {
    Param(
        [parameter(Mandatory=$true)]
        [string]$cumulativeFileURL,
        [parameter(Mandatory=$true)]
        [string]$bitDefenderPath
    )
    try {
        $zipFilePath = Join-Path $bitDefenderPath ([System.IO.Path]::GetFileName($cumulativeFileURL))
        Download-File $cumulativeFileURL $zipFilePath
        Extract-Zip $zipFilePath $bitDefenderPath
        Write-Host "アップデートが完了しました"
    } catch {
        Write-Warning "アップデートに失敗しました"
        Write-Debug ($_.Exception.tostring() + $_.InvocationInfo.PositionMessage)
    }
    try {
        if (Test-Path $zipFilePath) {
            Remove-Item $zipFilePath -Force
        }
        Get-ChildItem $bitDefenderPath | Where-Object{ $_.Name -match "BIT[0-9A-F]{4}\.tmp" } | ForEach-Object {
            Remove-Item $_.FullName -Force
        }
    } catch {
        Write-Warning "アップデート用ファイルの削除に失敗しました"
        Write-Debug ($_.Exception.tostring() + $_.InvocationInfo.PositionMessage)
    }
}

function Test-RegistryKeyValue {
    Param(
        [parameter(Mandatory=$true)]
        [string]$path,
        [parameter(Mandatory=$true)]
        [string]$name
    )
    if (!(Test-Path -Path $path -PathType Container)) {
        return $false
    }
    $properties = Get-ItemProperty -Path $path
    if (!$properties) {
        return $false
    }
    if (Get-Member -InputObject $properties -Name $name) {
        return $true
    } else {
        return $false
    }
}

function Get-InstallPath {
    if ((Get-WmiObject -Class Win32_OperatingSystem).OSArchitecture.Contains("64")) {
        $regPart = "Wow6432Node\"
    } else {
        $regPart = ""
    }
    $regKey = "HKLM:SOFTWARE\$($regPart)Microsoft\Windows\CurrentVersion\Uninstall\UiUicyBitDefCmd2_is1\"
    $regValueName = "InstallLocation"
    if (Test-RegistryKeyValue $regKey $regValueName) {
        return (Get-ItemProperty -Path $regKey).$regValueName
    } else {
        return $null
    }
}

function Pause {
    if (!$psISE) {
        Write-Host "続行するには何かキーを押してください..." -NoNewLine
        [void][System.Console]::ReadKey($true)
    }
}

$installPath = Get-InstallPath
if (!$installPath) {
    $installPath = $PSScriptRoot
}
$bdcExePath = Join-Path $installPath "\bdc.exe"
$proctermExePath = Join-Path $installPath "\procterm.exe"
$updateTxtPath = Join-Path $installPath "\Plugins\update.txt"
$batName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

if (Test-Path $updateTxtPath) {
    foreach($fileTxt in Get-Content $updateTxtPath) {
        if ($fileTxt -match "^Update time GMT: (?<unixTime>\d*)$") {
            break
        }
    }
    $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
    $updateTime = $origin.AddSeconds($matches.unixTime)
    if (Test-Update $updateFileURL $updateTime) {
        Update-Plugins $updateFileURL $installPath
    }
} elseif (Test-Update $updateFileURL ([DateTime]::MinValue)) {
    Update-Plugins $updateFileURL $installPath
} else {
    Write-Warning "アップデートに失敗しました"
}

Write-Host
if (Test-Path $updateTxtPath) {
    Get-Content $updateTxtPath
}
Write-Host
if (($batName -eq "bdcscan") -and (Test-Path $bdcExePath)) {
    &$bdcExePath $args /arc /mail /prompt /move /moves /infp=Infected /susp=Suspected /list | Select-Object -Skip 3
    Write-Host "ウィルススキャンの結果は以上です"
}
if (!$silent) {
    Pause
}
if (Test-Path $proctermExePath) {
    &$proctermExePath conime.exe
}