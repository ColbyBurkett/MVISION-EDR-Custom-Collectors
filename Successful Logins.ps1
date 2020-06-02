# This collector returns the Successful Logins from Windows Event Log
$OutputEncoding = New-Object -typename System.Text.UTF8Encoding
# resize PS buffer size in order to avoid undesired line endings or trims in the output
$pshost = get-host
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 3000
$pswindow.buffersize = $newsize

$events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID="4624"} -MaxEvents 300
foreach ($event in $events)
{
    $eventXML = [xml]$event.ToXml()
    $eventArray = New-Object -TypeName PSObject -Property @{


    EventID = $event.id
    EventTime = $event.timecreated
    UserName = $eventXML.Event.EventData.Data[5].'#text'
    Domain = $eventXML.Event.EventData.Data[6].'#text'
    LogonType = $eventXML.Event.EventData.Data[8].'#text'
    NetworkInformation = $eventXML.Event.EventData.Data[18].'#text'
    }
    $eventid=$eventarray.eventid
    $eventuser=$eventarray.username
    $eventdomain=$eventarray.domain
    $starttype=$eventarray.Starttype
    [datetime]$eventtime=$eventarray.eventtime
    [string]$dateformat = 'yyyy-MM-dd HH:mm:ss'
    $finaltime = $eventtime.ToString($dateformat)
    $eventlogontype=$eventarray.logontype
    $sourceIP=$eventArray.Networkinformation


 
    write-output "$eventuser,$eventDomain,$eventlogontype,$finaltime,$sourceIP"
}
