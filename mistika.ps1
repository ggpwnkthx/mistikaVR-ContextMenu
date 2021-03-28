param(
    [string]$InputPath
)
$dir_scope = "$env:APPDATA\Eclatech"
$reg_scope = "HKCU:\SOFTWARE\Eclatech"

function Install-ContextMenus() {
    if (!(Test-Path $dir_scope)) {
        New-Item -ItemType Directory -Force -Path $dir_scope | Out-Null
    }
    if (!(Test-Path $reg_scope)) {
        New-Item -Path "HKCU:\SOFTWARE" -Name "Eclatech" | Out-Null
    }
    
    # Script placement
    $self_install_path = "$dir_scope\Scripts"
    $self_install_path += "\"
    $self_path = New-Object System.IO.FileInfo($PSCommandPath)
    $self_install_path += $self_path.BaseName
    $self_install_path += $self_path.Extension

    # Context menu
    $reg_context = "HKCU:\SOFTWARE\Classes\Directory\Background\shell"
    if (!(Test-Path $reg_context)) {
        New-Item -Path $reg_context
    }

    $reg_context_eclatech = "$reg_context\eclatech"
    if (!(Test-Path $reg_context_eclatech)) {
        New-Item -Path $reg_context_eclatech | Out-Null
    }
    if (!(Get-ItemProperty -Path $reg_context_eclatech).PSObject.Properties.Name -contains "SubCommands") {
        Get-Item -Path $reg_context_eclatech | New-ItemProperty -Name "SubCommands" -Value ""
    }
    if (!(Get-ItemProperty -Path $reg_context_eclatech).PSObject.Properties.Name -contains "MUIVerb") {
        Get-Item -Path $reg_context_eclatech | New-ItemProperty -Name "MUIVerb" -Value "Eclatech"
    }

    $reg_context_eclatech_shell = "$reg_context_eclatech\shell"
    if (!(Test-Path $reg_context_eclatech_shell)) {
        New-Item -Path $reg_context_eclatech_shell | Out-Null
    }

    $reg_context_eclatech_shell_mistika_vr = "$reg_context_eclatech_shell\mistika_vr"
    if (!(Test-Path $reg_context_eclatech_shell_mistika_vr)) {
        New-Item -Path $reg_context_eclatech_shell_mistika_vr | Out-Null
    }
    if (!(Get-ItemProperty -Path $reg_context_eclatech_shell_mistika_vr).PSObject.Properties.Name -contains "MUIVerb") {
        Get-Item -Path $reg_context_eclatech_shell_mistika_vr | New-ItemProperty -Name "MUIVerb" -Value "mistika VR - Render All"
    }
    $reg_context_eclatech_shell_mistika_vr_command = "$reg_context_eclatech_shell_mistika_vr\command"
    if (!(Test-Path $reg_context_eclatech_shell_mistika_vr_command)) {
        New-Item -Path $reg_context_eclatech_shell_mistika_vr_command | Out-Null
    }
    Set-ItemProperty -Path $reg_context_eclatech_shell_mistika_vr_command -Name "(Default)" -Value "PowerShell.exe -ExecutionPolicy Bypass -NoExit -Command `"& '$self_install_path' -InputPath '%v'`""
}

# Github releases functions
function Get-GithubRelease {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    Param(
        [Parameter(
            Mandatory = $true,
            Position = 0
        )]
        [string] $Repo,
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ParameterSetName = "Prereleases"
        )]
        [switch] $Prereleases,
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ParameterSetName = "Latest"
        )]
        [switch] $Latest,
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ParameterSetName = "Name"
        )]
        [switch] $Name,
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ParameterSetName = "Tag"
        )]
        [switch] $Tag
    )
    $URI = "https://api.github.com/repos/$repo/releases"
    if ($Latest) {
        $URI += "/latest"
    }
    $releases = Invoke-RestMethod -Method GET -Uri $URI
    if (!$Prereleases) {
        $releases = $releases | Where-Object { $_.prerelease -eq $false }
    }
    if ($Name) {
        return $releases | Where-Object { $_.name -eq $Name }
    }
    if ($Tag) {
        return $releases | Where-Object { $_.tag_name -eq $Tag }
    }
    return , ($releases | Sort-Object published_at -Descending)
}

function Get-GithubAsset {
    Param(
        [Parameter(ValueFromPipeline, Mandatory = $true)] $Release,
        [Parameter(Mandatory = $false)] [switch] $Relative
    )
    $assets = @()
    $Release.assets | ForEach-Object { $assets += $_ }
    if ($Relative) {
        if ($assets.Length -gt 1) {
            switch ($ENV:PROCESSOR_ARCHITECTURE) {
                "X86" {
                    $asset = $assets | Where-Object { 
                        $_.name -like "*x86*"
                    }
                    if ($asset.Length -gt 0) { $assets = $asset }
                }
                "AMD64" {
                    $asset = $assets | Where-Object { 
                        $_.name -like "*x64*" -or
                        $_.name -like "*_64*"
                    }
                    if ($asset.Length -gt 0) { $assets = $asset }
                }
                "ARM32" {
                    $asset = $assets | Where-Object { 
                        $_.name -like "*arm32*" -or
                        (
                            $_.name -like "*arm*" -and
                            $_.name -notlike "*arm64*"
                        )
                    }
                    if ($asset.Length -gt 0) { $assets = $asset }
                }
                "ARM64" {
                    $asset = $assets | Where-Object { 
                        $_.name -like "*arm64*"
                    }
                    if ($asset.Length -gt 0) { $assets = $asset }
                }
            }
            
        }
        if ($assets.Length -gt 1) {
            if ($IsWindows -or $ENV:OS) {
                $asset = $assets | Where-Object { 
                    $_.name -like "*win*"
                }
                if ($asset.Length -gt 0) { $assets = $asset }
            }
            if ($IsMacOS) {
                $asset = $assets | Where-Object { 
                    $_.name -like "*mac*" -or
                    $_.name -like "*osx*"
                }
                if ($asset.Length -gt 0) { $assets = $asset }
            }
            if ($IsLinux) {
                $asset = $assets | Where-Object { 
                    $_.name -like "*linux*"
                }
                if ($asset.Length -gt 0) { $assets = $asset }
            }
            
        }
        if ($assets.Length -gt 1) {
            return $assets | Sort-Object download_count -Descending | Select-Object -First 1
        }
        else {
            return $assets
        }
    }
    else {
        return $assets
    }
}

function Download-GithubAsset {
    Param(
        [Parameter(ValueFromPipeline, Mandatory = $true)] $Asset
    )
    $FileName = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $(Split-Path -Path $Asset.browser_download_url -Leaf)
    Invoke-WebRequest -Uri $Asset.browser_download_url -Out $FileName    
    switch ($([System.IO.Path]::GetExtension($Asset.browser_download_url))) {
        ".zip" {
            $tempExtract = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $((New-Guid).Guid)
            Expand-Archive -Path $FileName -DestinationPath $tempExtract -Force
            Remove-Item $FileName -Force
            return $tempExtract
        }
        default {
            return $FileName
        }
    }
}

function Self-Upgrade ([string]$InputPath) {
    Install-ContextMenus
    $current_publish = (New-TimeSpan -Start (Get-Date -Date "01/01/1970") -End (Get-ChildItem $PSCommandPath).LastWriteTime).TotalSeconds
    $release = Get-GithubRelease -Repo ggpwnkthx/mistikaVR-ContextMenu -Latest
    $release_publish = (New-TimeSpan -Start (Get-Date -Date "01/01/1970") -End (Get-Date -Date $release.published_at)).TotalSeconds
    if ($release_publish -gt $current_publish) {
        Write-Host "Performing self-update..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-Module -Name PackageManagement -Force -MinimumVersion 1.4.6 -Scope CurrentUser -AllowClobber -Repository PSGallery | Out-Null
        Register-PackageSource -Name nuget.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet | Out-Null
        Set-PackageSource -Name nuget.org -Trusted | Out-Null
        Copy-Item -Path ($release | Get-GithubAsset | Download-GithubAsset) -Destination $PSCommandPath -Force
        & $PSCommandPath -InputPath $InputPath
        exit
    }
}

# Self-Install
if ($PSScriptRoot -ne "$dir_scope\Scripts") {
    # Registry
    Install-ContextMenus

    # Script placement
    $self_install_path = "$dir_scope\Scripts"
    if (!(Test-Path $self_install_path)) {
        New-Item -ItemType Directory -Force -Path $self_install_path | Out-Null
    }
    Copy-Item $PSCommandPath -Destination $self_install_path -Force

    Write-Host -NoNewLine 'Installed successfully! Use folder context menus to run the service.';
}
else {
    if ($InputPath -eq $null) {
        $InputPath = Get-Folder
    }
    Self-Upgrade -InputPath $InputPath

    Get-ChildItem -Path $InputPath -Filter "*.rnd" | ForEach-Object { 
        Start-Process -FilePath 'C:\Program Files\SGO Apps\Mistika VR\bin\vr' -ArgumentList "-r", $_.FullName -NoNewWindow -Wait 
    }
}