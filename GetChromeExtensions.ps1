# Based on script on SpiceWorks by "Bellows"
# Changes:
#     Refactored for accuracy and brevity
#     Added another location of extension name
#     Changed output to on-screen CSV
# Returns all versions of Extensions on the system, active or not
# Currently English only
$OutputEncoding = New-Object -typename System.Text.UTF8Encoding
# resize PS buffer size in order to avoid undesired line endings or trims in the output
$pshost = get-host
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 3000
$pswindow.buffersize = $newsize

##: The extensions folder is in local appdata 
$extension_folders = Get-ChildItem -Path "C:\users\*\appdata\local\Google\Chrome\User Data\Default\Extensions\*" -ErrorAction SilentlyContinue

##: Loop through each extension folder
foreach ($extension_folder in $extension_folders ) {

    ##: Get the version specific folder within this extension folder
    $version_folders = Get-ChildItem -Path "$($extension_folder.FullName)"

    ##: Loop through the version folders found
    foreach ($version_folder in $version_folders) {
        ##: The extension folder name is the app id in the Chrome web store
        $appid = $extension_folder.BaseName
        $fullpath = ""

        ##: First check the manifest for a name
        $name = ""
        if( (Test-Path -Path "$($version_folder.FullName)\manifest.json") ) {
            try {
                $json = Get-Content -Raw -Path "$($version_folder.FullName)\manifest.json" | ConvertFrom-Json
                $name = $json.name
            }
            catch {
                $name = ""
            }
        }
        ##: If we find _MSG_ in the manifest it's probably an app
        if( $name -like "*_MSG_*" ) {
            ##: Sometimes the folder is en, sometimes in en_US
            if( Test-Path -Path "$($version_folder.FullName)\_locales\en\messages.json" ) {
                $localePath = "$($version_folder.FullName)\_locales\en\messages.json"
                }
            elseif( !$fullpath -and (Test-Path -Path "$($version_folder.FullName)\_locales\en_US\messages.json") ) {
                $localePath = "$($version_folder.FullName)\_locales\en_US\messages.json"
                }
            try { 
                $json = Get-Content -Raw -Path $localePath | ConvertFrom-Json
                $name = $json.appName.message
                ##: Try different ways to get the Extension name
                if(!$name) {
                    $name = $json.extName.message
                    }
                if(!$name) {
                    $name = $json.ext_Name.message
                    }
                if(!$name) {
                    $name = $json.extensionName.message
                    }
                if(!$name) {
                    $name = $json.app_name.message
                    }
                if(!$name) {
                    $name = $json.application_title.message
                    }
                }
            catch { 
                    $name = ""
                }
            }
        ##: If we can't get a name from the extension use the app id instead
        if( !$name ) {
            $name = $appid
        }
        if ($version_folders) {
            Write-Host ("""{0}"",""{1}"",""{2}""" -f $name,$version_folder,$appid)
            }
        }
}