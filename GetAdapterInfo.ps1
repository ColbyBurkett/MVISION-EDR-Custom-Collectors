# This collector returns the IP Address and Name for Network Interfaces, including VPN
$OutputEncoding = New-Object -typename System.Text.UTF8Encoding
# resize PS buffer size in order to avoid undesired line endings or trims in the output
$pshost = get-host
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 3000
$pswindow.buffersize = $newsize

$netConnections = Get-NetIPConfiguration | Where-Object {$_.NetAdapter.Status -ne 'Disconnected'}
foreach ($row in $netConnections) {
    foreach ($IPAddress in $row.IPv4Address) {
        Write-Host ("""{0}"",""{1}""" -f $IPAddress.IPAddress,$row.InterfaceDescription)
    }
}