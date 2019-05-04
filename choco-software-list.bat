@powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
REM ### enable WSL
@powershell -NoProfile -ExecutionPolicy Bypass -Command "Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -NoRestart"
REM ### system
choco install -y 7zip.install
choco install -y notepadplusplus.install
REM # configure 7-zip options
reg add HKCU\Software\7-Zip\FM\ /v AlternativeSelection /t REG_DWORD /d 0 /f
reg add HKCU\Software\7-Zip\FM\ /v FlatViewArc0 /t REG_DWORD /d 0 /f
reg add HKCU\Software\7-Zip\FM\ /v FlatViewArc1 /t REG_DWORD /d 0 /f
reg add HKCU\Software\7-Zip\FM\ /v FullRow /t REG_DWORD /d 1 /f
reg add HKCU\Software\7-Zip\FM\ /v ListMode /t REG_DWORD /d 771 /f
reg add HKCU\Software\7-Zip\FM\ /v ShowDots /t REG_DWORD /d 1 /f
reg add HKCU\Software\7-Zip\FM\ /v ShowGrid /t REG_DWORD /d 1 /f
reg add HKCU\Software\7-Zip\FM\ /v ShowRealFileIcons /t REG_DWORD /d 1 /f
reg add HKCU\Software\7-Zip\FM\ /v ShowSystemMenu /t REG_DWORD /d 0 /f
reg add HKCU\Software\7-Zip\FM\ /v SingleClick /t REG_DWORD /d 0 /f
reg add HKCU\Software\7-Zip\FM\ /v Viewer /t REG_SZ /d "C:\Program Files\Notepad++\notepad++.exe" /f
reg add HKCU\Software\7-Zip\FM\ /v Editor /t REG_SZ /d "C:\Program Files\Notepad++\notepad++.exe" /f
reg add HKCU\Software\7-Zip\Options\ /v CascadedMenu /t REG_DWORD /d 0 /f
reg add HKCU\Software\7-Zip\Options\ /v ContextMenu /t REG_DWORD /d 4902 /f
choco install -y sysinternals
choco install -y processhacker
choco install -y firefox
choco install -y adblockplus-firefox
choco install -y sumatrapdf
choco install -y ag
choco install -y hxd
choco install -y jq
choco install -y dnspy
choco install -y curl
choco install -y gow
choco install -y gpg4win-vanilla
choco install -y bleachbit
choco install -y winscp
choco install -y putty
choco install -y conemu
REM choco install -y nirlauncher
choco install -y vlc
REM choco install -y greenshot
REM ### dev related
choco install -y powershell
choco install -y git
choco install -y winmerge
choco install -y x64dbg.portable
choco install -y ghidra
REM choco install -y agentransack
choco install -y openssh
choco install -y bitnami-xampp
choco install -y apimonitor
choco install -y heidisql
choco install -y windbg
choco install -y golang
choco install -y lazarus
choco install -y strawberryperl
choco install -y python2
choco install -y vcpython27
choco install -y pywin32
choco install -y visualstudiocode
REM choco install -y visualstudio2017community
REM choco install -y azure-cli
REM choco install -y windowsazurepowershell 
choco install -y awscli
choco install -y awstools.powershell
REM choco install -y vagrant
choco install -y everything
REM ### other
choco install -y vcredist2015 
choco install -y vcredist2013
choco install -y vcredist2010
choco install -y vcredist2008
choco install -y jre8
choco install -y burp-suite-free-edition
choco install -y fiddler4
choco install -y postman
psshutdown -accepteula -f -r -t 60
