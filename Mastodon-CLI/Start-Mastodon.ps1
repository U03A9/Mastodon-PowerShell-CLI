<# 

This is the main launch script that will import & launch the Mastodon command line engine.

This will serve as the 'front-end' launcher until a compiled .exe is created.

#>

function Get-ScriptDirectory
{
    $Global:ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
}

Function Run-Mastodon{


Import-Module "$ScriptDir\Mastodon-CLI.psm1" | Out-File -FilePath $ScriptDir\debug.txt
Start-Mastodon

}

Get-ScriptDirectory

Run-Mastodon