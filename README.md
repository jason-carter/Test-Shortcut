# Test-Shortcut
PowerShell script to test the shortcuts of a passed folder are valid

Example usage, which is included in the script (so please change before running), is as follows:

```PowerShell
Test-Shortcut -path "${en:USERPROFILE}\AppData\Roaming\Microsoft\Windows\Recent" |
    Where-Object { $_.IsWorking -eq $false} |
    Where-Object { $_.Location.Length -gt 0 } |
    Where-Object { $_.Location.StartsWith('I:\') -or $_.Location.StartsWith('H:\') } |
    ForEach-Object { Remove-Item $_.ShortcutFullName }
```

This passes the path of the Recents folder to the `Test-Shortcut` script, which lists all shortcuts in that folder, along with whether they are working or not, and then filters out those that are not working, then those with a non-empty shortcut location, then filters our the remaining shortcuts that point tothe I or H drive, and finally removes that shortcut.
