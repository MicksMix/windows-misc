rem @powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
refreshenv
REM # system
choco install -y 7zip.install
choco install -y notepadplusplus.install
choco install -y firefox
choco install -y adblockplus-firefox
choco install -y sumatrapdf
choco install -y ag
choco install -y classic-shell
choco install -y hxd
choco install -y jq
choco install -y ilspy
choco install -y curl
choco install -y gow
choco install -y gpg4win-vanilla
choco install -y bleachbit
choco install -y sysinternals
choco install -y processhacker
choco install -y winscp
choco install -y putty
choco install -y conemu
choco install -y nirlauncher
choco install -y vlc
choco install -y greenshot
REM # dev related
choco install -y git
choco install -y winmerge
choco install -y agentransack
choco install -y openssh
choco install -y heidisql
choco install -y windbg
choco install -y golang
choco install -y strawberryperl
choco install -y python2
choco install -y vcpython27
choco install -y pywin32
choco install -y visualstudiocode
choco install -y visualstudio2015community
choco install -y everything
REM # other
choco install -y jre8
choco install -y burp-suite-free-edition
choco install -y fiddler4
