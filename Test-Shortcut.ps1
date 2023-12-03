Function Test-Shortcut
{
    [CmdletBinding()]
	Param (
        [Parameter(Mandatory=$True, HelpMessage="Please enter Path of Shortcut, Example c:\folder", Position=1)]
        [string] $path
    )

    $ErrorActionPreference =' Silentlycontinue'

    $Shortcuts = Get-ChildItem -Recurse $path -Include *.lnk -Force
    $Shell     = New-Object -ComObject WScript.Shell

    foreach ($Shortcut in $Shortcuts) {
        $checks = $shell.CreateShortcut($Shortcut).TargetPath   
     
        foreach($check in $checks) {
            [Pscustomobject]@{ 
                IsWorking        = Test-Path $($check)
                Shortcut         = $Shortcut.name
                ShortcutFullName = $Shortcut.FullName
                Location         = $($check)
            }
        }
    }
  }

  # Remove all broken links in the recents folder that point to the I or H drive
Test-Shortcut -path "${env:USERPROFILE}\AppData\Roaming\Microsoft\Windows\Recent" |
    Where-Object { $_.IsWorking -eq $false} |
    Where-Object { $_.Location.Length -gt 0 } |
    Where-Object { $_.Location.StartsWith('I:\') -or $_.Location.StartsWith('H:\') } |
    Select-Object ShortcutFullName |
    ForEach-Object { Remove-Item $_.ShortcutFullName }
