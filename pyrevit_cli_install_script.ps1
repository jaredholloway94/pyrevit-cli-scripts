function Write-HostQuiet
{
    [CmdletBinding()]

    param
    (
        [Parameter()]
        [string]
        $Message,

        [Parameter()]
        [switch]
        [Alias("q")]
        $Quiet
    )

    if (-not $Quiet) {Write-Host $Message}
}


function Test-pyRevitCliVersion
{
    try
    {
        $pyrevit_cli_version = pyrevit -V 2>&1

        if ($pyrevit_cli_version[0].Exception.Message -eq 'You must install or update .NET to run this application.')
        {
            # pyRevitCLI is installed, but .NET 8.0 is not.
            return -1
        }
        elseif ($pyrevit_cli_version[0] -like 'pyrevit *')
        {
            # pyRevitCLI is installed.
            return $pyrevit_cli_version[0].split(' +')[1].substring(1)
        }
    }
    catch
    {
        # pyRevitCLI is not installed.
        return 0
    }
       
}


function Remove-pyRevitPreviousInstalls
{
    [CmdletBinding()]

    param
    (
        [Parameter()]
        [switch]
        $Quiet
    )

    # paths
    $cli_unins_path = "C:\Program Files\pyRevit-Master\unins000.exe"
    $sys_unins_path = "C:\Program Files\pyRevit CLI\unins000.exe"
    $sys_config_path = "C:\ProgramData\pyRevit"

    # uninstall pyRevit CLI
    if (Test-Path $cli_unins_path)
    {
        Write-HostQuiet "Removing pyRevit CLI ..." -Quiet:$Quiet
        Start-Process -FilePath $cli_unins_path `
            -ArgumentList "/VERYSILENT /NORESTART" `
            -Wait `
            -NoNewWindow
        Write-HostQuiet "Done.`n" -Quiet:$Quiet
    }

    # uninstall pyRevit SYSTEM app
    if (Test-Path $sys_unins_path)
    {
        Write-HostQuiet "Removing system pyRevit install ..." -Quiet:$Quiet
        Start-Process -FilePath $sys_unins_path `
            -ArgumentList "/VERYSILENT /NORESTART" `
            -Wait `
            -NoNewWindow
        Write-HostQuiet "Done.`n" -Quiet:$Quiet
    }

    if (Test-Path $sys_config_path)
    {
        Write-HostQuiet "Removing system pyRevit_config.ini ..." -Quiet:$Quiet
        Get-ChildItem $sys_config_path | Remove-Item -Force
        Remove-Item $sys_config_path -Force
        Write-HostQuiet "Done.`n" -Quiet:$Quiet
    }

    # uninstall pyRevit USER apps
    Get-ChildItem "C:\Users" | ForEach-Object {

        $usr_unins_path = "C:\Users\"+$($_.name)+"\AppData\Roaming\pyRevit-Master\unins000.exe"
        $usr_config_path = "C:\Users\"+$($_.name)+"\AppData\Roaming\pyRevit"

        if (Test-Path $usr_unins_path)
        {
            Write-HostQuiet "Removing user pyRevit install for $($_.name)..." -Quiet:$Quiet
            Start-Process -FilePath $usr_unins_path `
                -ArgumentList "/VERYSILENT" `
                -Wait `
                -NoNewWindow
            Remove-Item $usr_config_path -Recurse -Force
            Write-HostQuiet "Done.`n" -Quiet:$Quiet
        }

        if (Test-Path $usr_config_path)
        {
            Write-HostQuiet "Removing user pyrevit_config.ini for $($_.name)..." -Quiet:$Quiet
            Get-ChildItem $usr_config_path | Remove-Item -Force
            Remove-Item $usr_config_path -Force
            Write-HostQuiet "Done.`n" -Quiet:$Quiet
        }
    }

    Write-HostQuiet "All previous pyRevit installs have been removed." -Quiet:$Quiet
}


function Update-pyRevitCliVersion
{
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory=$True)]
        [string]
        $DotNetInstallerPath,

        [Parameter(Mandatory=$True)]
        [string]
        $PyRevitCliInstallerPath,

        [Parameter()]
        [int]
        $RecursionDepth,

        [Parameter()]
        [switch]
        [Alias("q")]
        $Quiet,

        [Parameter()]
        [switch]
        [Alias("c")]
        $Clean
    )

    # prevent infinite recusion on installer failure
    if ($RecursionDepth -gt 3)
    {
        Throw "Recursion depth limit reached."
    }

    # if desired, remove all previous installs of pyRevit
    if ($Clean)
    {
        Remove-pyRevitPreviousInstalls -Quiet:$Quiet
    }

    # ensure local folder exists to copy installer from (potential) network location to local location
    # (this is the only way to guarantee security dialog will not pop up during silent install)
    if (-not (Test-Path "C:\pyRevit"))
    {
        mkdir "C:\pyRevit" > $null
    }

    # main
    switch ((Test-pyRevitCliVersion)[0])
    {
        -1 
        {  
            # pyRevitCLI is installed, but .NET Framework is not.
            # install .NET Framework from cached installer.

            Write-HostQuiet "Installing .NET Framework ..." -Quiet:$Quiet

            $LocalDotNetInstallerPath = Join-Path -Path "C:\pyRevit" `
                -ChildPath:([System.IO.Path]::GetFileName($DotNetInstallerPath))

            Copy-Item $DotNetInstallerPath $LocalDotNetInstallerPath

            Start-Process -FilePath:$LocalDotNetInstallerPath `
                -ArgumentList:"/q /norestart" `
                -Wait `
                -NoNewWindow

            Write-HostQuiet "Done.`n" -Quiet:$Quiet

            Update-pyRevitCliVersion `
                -DotNetInstallerPath:$DotNetInstallerPath `
                -PyRevitCliInstallerPath:$PyRevitCliInstallerPath `
                -RecursionDepth:($RecursionDepth+1) `
                -Quiet:$Quiet
        }

        0
        {   
            # pyRevit CLI is not installed.
            # install pyRevitCLI 5.0 from cached installer.

            Write-HostQuiet "Installing pyRevit CLI v5 ..." -Quiet:$Quiet

            $LocalPyRevitCliInstallerPath = Join-Path -Path:"C:\pyRevit" `
                -ChildPath:([System.IO.Path]::GetFileName($PyRevitCliInstallerPath))

            Copy-Item $PyRevitCliInstallerPath $LocalPyRevitCliInstallerPath

            Start-Process -FilePath:$LocalPyRevitCliInstallerPath `
                -ArgumentList:"/VERYSILENT /NORESTART" `
                -Wait `
                -NoNewWindow

            Write-HostQuiet "Done.`n" -Quiet:$Quiet

            Update-pyRevitCliVersion `
                -DotNetInstallerPath:$DotNetInstallerPath `
                -PyRevitCliInstallerPath:$PyRevitCliInstallerPath `
                -RecursionDepth:($RecursionDepth+1) `
                -Quiet:$Quiet
        }

        4
        {   
            # pyRevit CLI is installed, but out-of-date.
            # install pyRevitCLI 5.0 from cached installer.

            Write-HostQuiet "Updating pyRevit CLI to v5 ..." -Quiet:$Quiet

            $LocalPyRevitCliInstallerPath = Join-Path -Path:"C:\pyRevit" `
                -ChildPath:([System.IO.Path]::GetFileName($PyRevitCliInstallerPath))

            Copy-Item $PyRevitCliInstallerPath $LocalPyRevitCliInstallerPath

            Start-Process -FilePath:$LocalPyRevitCliInstallerPath `
                -ArgumentList:"/VERYSILENT /NORESTART" `
                -Wait `
                -NoNewWindow

            Write-HostQuiet "Done.`n" -Quiet:$Quiet

            Update-pyRevitCliVersion `
                -DotNetInstallerPath:$DotNetInstallerPath `
                -PyRevitCliInstallerPath:$PyRevitCliInstallerPath `
                -RecursionDepth:($RecursionDepth+1) `
                -Quiet:$Quiet
        }
        
        5
        {   
            # pyRevit CLI is installed, and up-to-date.
            # continue.

            Write-HostQuiet "PyRevit CLI and its dependencies are up to date." -Quiet:$Quiet
        }
    }
}
