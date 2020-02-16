# adapted from: https://johnmj.wordpress.com/2010/11/11/sharepoint-2010-and-localhost-2/
param (
    [parameter (ParameterSetName='file_input', Mandatory=$true, Position=0)]
    [string] $file,
    [parameter (ParameterSetName='pipe_input', Mandatory=$true, ValueFromPipeline=$true)]
    [string] $lines,
    [parameter (Mandatory=$false)]
    [string] $dir = $env:APPDATA + "\Microsoft\windows\Start Menu\Programs\HotKeys\”
)



function Create-Shortcut {
    [CmdletBinding()]
    param (
        [parameter (Position=0,Mandatory=$true)]
        [string[]]$path,
        [parameter (Position=1,Mandatory=$true)]
        [string] $ShortcutKeys,
        [parameter (Position=2,Mandatory=$false)]
        [string] $ShortcutName
    )
    
    $shell = New-Object -comObject wscript.Shell
    if ($ShortcutName) {
        $f = $ShortcutName + ".lnk"
    } else {
        $f = [guid]::NewGuid().toString() + ".lnk"
    }
    $sc = $shell.createshortcut($dir + $f)
    $sc.TargetPath = "$path"
    if ($shortcutKeys.TocharArray().Length -eq 1) { $ShortcutKeys = "CTRL+ALT+" + $ShortcutKeys }
    $sc.Hotkey = $ShortcutKeys
    write-host $sc.save()
    write-Host "new shortcut '$f' created for '$path'”
}



write-Host "Removing previous shortcuts...”
Remove-Item -force -Recurse -ErrorAction Ignore -path $dir
mkdir $dir >$null

write-Host "Parsing new shortcuts list...”
if ($line) {
    $lines = @($line)
} else {
    [Array] $lines = Get-Content $file
}

write-Host "Creating Shortcuts...”
foreach ($line in $lines) {
    $fields = $line.split(”/”, [System.stringSplitoptions]::RemoveemptyEntries)
    create-shortcut @fields
}

write-Host "`nAll done.”
dir $dir
