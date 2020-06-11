<#
.SYNOPSIS
  Dumps all install Internet Explorer Extensions (add-ons)
.DESCRIPTION
  Walks the registry to identify Internet Explorer Extensions for the currently logged in user
.NOTES
  Version:        0.1
  Author:         @brad_anton (OpenDNS)
  Creation Date:  June 1, 2016
  Purpose/Change: Initial script development
  Modifications by Colby Burkett:
                           CSV Output
                           Continue on Errors
                           General improvements to results
  
.EXAMPLE
"Evernote extension","C:\Program Files (x86)\Evernote\Evernote\EvernoteIEx64.dll"
"HTML Document","C:\Windows\System32\mshtml.dll"
"Java(tm) Plug-In 2 SSV Helper","C:\Program Files\Java\jre1.8.0_241\bin\jp2ssv.dll"
"Java(tm) Plug-In SSV Helper","C:\Program Files\Java\jre1.8.0_241\bin\ssv.dll"
#>

<#
    Utility Functions
#>

$OutputEncoding = New-Object -typename System.Text.UTF8Encoding
# resize PS buffer size in order to avoid undesired line endings or trims in the output
$pshost = get-host
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 3000
$pswindow.buffersize = $newsize

function Lookup-Clsid
{
    Param([string]$clsid)
    $CLSID_KEY = 'HKLM:\SOFTWARE\Classes\CLSID'

    If ( Test-Path $CLSID_KEY\$clsid) {
            $name = (Get-ItemProperty -ErrorAction SilentlyContinue -Path $CLSID_KEY\$clsid).'(default)'
            $dll = (Get-ItemProperty -ErrorAction SilentlyContinue -Path $CLSID_KEY\$clsid\InProcServer32).'(default)'
    }
    $name, $dll
}

function Make-Extension
{
    Param([string]$clsid, [string]$name, [string]$dll)
    
    $extension = New-Object PSObject -Prop (@{'CLSID' = $clsid;
                    'Name' = $name;
                    'DLL' = $dll })

    $extension
}

# Resulting list of Extension Properties 
$extensions = @()

<# 
    Extensions are identified in these Keys as properties containing the 
    CLSID.
#> 

$registry_keys = @( 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects',
                    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects',
                    'HKCU:\Software\Microsoft\Internet Explorer\UrlSearchHooks',
                    'HKLM:\Software\Microsoft\Internet Explorer\Toolbar',
                    'HKLM:\Software\Wow6432Node\Microsoft\Internet Explorer\Toolbar',
                    'HKCU:\Software\Microsoft\Internet Explorer\Explorer Bars',
                    'HKLM:\Software\Microsoft\Internet Explorer\Explorer Bars',
                    'HKCU:\Software\Wow6432Node\Microsoft\Internet Explorer\Explorer Bars',
                    'HKLM:\Software\Wow6432Node\Microsoft\Internet Explorer\Explorer Bars'
    )

ForEach ($key in $registry_keys) {
    If (Test-Path $key ) {
        $clsids = Get-Item -Path $key | Select-Object -Property Property | ForEach-Object Property
        ForEach ( $clsid in $clsids ) 
        {
            $name, $dll = Lookup-Clsid $clsid
            $extension = Make-Extension $clsid $name $dll
            $extensions += $extension
        }
    }
}

<# 
    Extensions are identified in these keys as subkeys named as the CLSID
#> 

$registry_keys = @( 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Ext\Stats' )

ForEach ($key in $registry_keys) {
    If (Test-Path $key ) {
        $clsids = Get-ChildItem $key -Name 
        ForEach ( $clsid in $clsids ) 
        {
            $name, $dll = Lookup-Clsid $clsid
            $extension = Make-Extension $clsid $name $dll
            $extensions += $extension
        }
    }
}

<#
    Extensions are identified in these keys as Values for the ClsidExtension
    Property within a subkeys named as some other ID.
#>

$registry_keys = @( 'HKCU:\Software\Microsoft\Internet Explorer\Extensions',
                    'HKLM:\Software\Microsoft\Internet Explorer\Extensions',
                    'HKCU:\Software\Wow6432Node\Microsoft\Internet Explorer\Extensions',
                    'HKLM:\Software\Wow6432Node\Microsoft\Internet Explorer\Extensions' )

ForEach ($key in $registry_keys) {
    If (Test-Path $key ) {
        $ids = Get-ChildItem $key -Name 
        ForEach ( $id in $ids ) 
        {
            $clsid = (Get-ItemProperty -ErrorAction SilentlyContinue -Path $key\$id -Name ClsidExtension).'ClsidExtension'
            $name, $dll = Lookup-Clsid $clsid
            $extension = Make-Extension $clsid $name $dll
            $extensions += $extension
        }
    }
}

$extensions = $extensions | Select-Object -Property Name, DLL | Sort-Object -Unique -Property Name, DLL

# Print CSV
ForEach ($extension in $extensions) {
    If ($extension.Name.Length -gt 0) {
        Write-Host ("""{0}"",""{1}""" -f $extension.Name,$extension.DLL)
    }
}