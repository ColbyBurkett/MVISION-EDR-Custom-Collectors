# This collector returns the Failed Logins from Windows Event Log
$OutputEncoding = New-Object -typename System.Text.UTF8Encoding
# resize PS buffer size in order to avoid undesired line endings or trims in the output
$pshost = get-host
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 3000
$pswindow.buffersize = $newsize

$events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4625} -MaxEvents 300
foreach ($event in $events)
{
    $eventXML = [xml]$event.ToXml()

    $eventArray = New-Object -TypeName PSObject -Property @{
    EventID = $event.id
    EventTime = $event.timecreated
    Domain = $eventXML.Event.EventData.Data[6].'#text'
    LogonType = $eventXML.Event.EventData.Data[10].'#text'
    failureReason = $eventXML.Event.EventData.Data[8].'#text'
    SubjectUserName = $eventXML.Event.EventData.Data[1].'#text'
    TargetUserName = $eventXML.Event.EventData.Data[5].'#text'
    NetworkInformation = $eventXML.Event.EventData.Data[19].'#text'
        }

    $eventid=$eventarray.eventid
    $eventuser=$eventarray.targetusername
    $eventdomain=$eventarray.domain
    $eventlogontype=$eventarray.logontype
    $failureReason = $eventArray.failureReason
    $targetUsername = $eventArray.SubjectUserName
    [datetime]$eventtime=$eventarray.eventtime
    [string]$dateformat = 'yyyy-MM-dd HH:mm:ss'
    $finaltime = $eventtime.ToString($dateformat)
    $sourceIP=$eventArray.Networkinformation

    write-output "$eventuser,$targetusername,$eventDomain,$eventlogontype,$failureReason,$finaltime,$sourceIP"
} 
