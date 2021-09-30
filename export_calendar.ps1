# @License: None (Large amounts of Stackoverflow code, credit belongs to the code segment's respective authors where it can be identified)

echo 'Beginning Outlook Calendar Export'
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$outlook = new-object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
$Appointments = $folder.Items
$Appointments.Sort("[Start]")
	
$sb = [System.Text.StringBuilder]::new()
# Fill in ICS/iCalendar properties based on RFC2445
[void]$sb.AppendLine("BEGIN:VCALENDAR")
[void]$sb.AppendLine("VERSION:2.0")
[void]$sb.AppendLine("METHOD:PUBLISH")
[void]$sb.AppendLine("PRODID:-//CHCC HIT//PowerShell ICS Migration Tool//EN")
[void]$sb.AppendLine("VERSION:2.0")
[void]$sb.AppendLine("CALSCALE:GREGORIAN")
[void]$sb.AppendLine("METHOD:REQUEST")
[void]$sb.AppendLine("BEGIN:VTIMEZONE")
[void]$sb.AppendLine("TZID:Pacific/Guam")
[void]$sb.AppendLine("X-LIC-LOCATION:Pacific/Guam")
[void]$sb.AppendLine("BEGIN:STANDARD")
[void]$sb.AppendLine("TZOFFSETFROM:+1000")
[void]$sb.AppendLine("TZOFFSETTO:+1000")
[void]$sb.AppendLine("TZNAME:ChST")
[void]$sb.AppendLine("DTSTART:19700101T000000")
[void]$sb.AppendLine("END:STANDARD")
[void]$sb.AppendLine("END:VTIMEZONE")

foreach($appt in ($Appointments)){
    echo 'Loading appointment...'
    #combining multiple calendar appointments is as simple as multiple BEGIN:VEVENT [...] END:VEVENT in the .ics (iCalendar) text document
    [void]$sb.AppendLine("BEGIN:VEVENT")
    [void]$sb.AppendLine("DTSTART;TZID=Pacific/Guam:" + "{0:yyyyMMddTHHmmss}" -f $x.Start) #InStartTimeZone) #Start is fine
    #[void]$sb.AppendLine("DTSTART:" + $x.StartInStartTimeZone)
    [void]$sb.AppendLine("DTEND;TZID=Pacific/Guam:" + "{0:yyyyMMddTHHmmss}" -f $x.End) #InEndTimeZone) #End is fine
	  #[void]$sb.AppendLine("RRULE:FREQ=WEEKLY;WKST=SU;UNTIL=20250101T135959Z;") #Needs correction of FREQ= value
    [void]$sb.AppendLine("DTSTAMP:20210929T010719Z")
    [void]$sb.AppendLine("ORGANIZER;CN=CNMI-CHCC ELC (Prod):" + $x.Organizer)
    [void]$sb.AppendLine("UID:" + [guid]::NewGuid())
	
    $x = $appt | Select-Object -Property *
    
    # Recover the Event Response
    $ParamResponseStatus = "NEEDS-ACTION"
    #$ParamResponseStatus = $x.ResponseStatus # (5) Recipient has not responded. (4) Meeting declined (3) Meeting accepted. (2) Meeting tentatively accepted.
    if($x.ResponseStatus -eq "2"){$ParamResponseStatus = "TENTATIVE"}
    if($x.ResponseStatus -eq "3"){$ParamResponseStatus = "CONFIRMED"}
    
    #some emails are aliased by name in RequiredAttendees and OptionalAttendees arrays and need a lookup table....
    $recipient_dict = @{}
    $recipient = $appt.Recipients
    foreach($r in ($recipient)){
                $recipient_dict.Add($r.Name,$r.Address)
                #echo $r.Name
                #echo $r.Address
                #$recipients_list.append($r)
                }
    try{
        foreach($attendee in ($x.RequiredAttendees.Split(";"))){
            if($recipient_dict.Get_Item($attendee) -like '*@*') {
                [void]$sb.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"REQ-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $recipient_dict.Get_Item($attendee) + ";X-NUM-GUESTS=0:mailto:" + $recipient_dict.Get_Item($attendee))
            }
            else{
                $attendeeStr = "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"REQ-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $attendee.Trim() + ";X-NUM-GUESTS=0:mailto:" + $attendee.Trim()
                if(-Not ($attendeeStr -like '*CN=;*')){[void]$sb.AppendLine($attendeeStr.TrimStart())} #dumb but effective
            }         
        }
        
    }
    catch{} #echo "0 Optional Attendees"}   
    #echo "Optional Attendees"
    #echo $x.OptionalAttendees.Length
    #if($x.OptionalAttendees.Length > 0){$recipient_dict.Get_Item($attendee)
    
    try{
        foreach($attendee in ($x.OptionalAttendees.Split(";"))){
            #echo $attendee
            if($recipient_dict.Get_Item($attendee) -like '*@*') {
                #echo $recipient_dict.Get_Item($attendee)
                [void]$sb.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"OPT-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $recipient_dict.Get_Item($attendee).Trim() + ";X-NUM-GUESTS=0:mailto:" + $recipient_dict.Get_Item($attendee).Trim())
            }
            else{
                #echo $recipient_dict.Get_Item($attendee)
                $attendeeStr = "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE="+"OPT-PARTICIPANT;PARTSTAT=" + $ParamResponseStatus + ";RSVP=TRUE;CN=" + $attendee.Trim() + ";X-NUM-GUESTS=0:mailto:" + $attendee.Trim()
                if(-Not ($attendeeStr -like '*CN=;*')){[void]$sb.AppendLine($attendeeStr.TrimStart())} #dumb but effective
            }
            #if($attendee -like '*@*') {
                #echo $attendee
            #    }
            #echo "69696969" 
            }
    }
    catch{} #echo "0 Optional Attendees"}   
	[void]$sb.AppendLine("CREATED:" + "{0:yyyyMMddTHHmmss}" -f $x.CreationTime)
	[void]$sb.AppendLine("DTSTAMP:" + "{0:yyyyMMddTHHmmss}" -f $x.CreationTime)
	[void]$sb.AppendLine("LAST-MODIFIED:" + [datetime]::$x.LastModificationTime)
	[void]$sb.AppendLine("LOCATION:" + $x.Location)
	[void]$sb.AppendLine("SEQUENCE:1")
	[void]$sb.AppendLine("STATUS" + $ParamResponseStatus)
	[void]$sb.AppendLine("SUMMARY:" + $x.Subject)
	[void]$sb.AppendLine("DESCRIPTION:" + $x.Body) # -Encoding UTF8 -Raw)
	[void]$sb.AppendLine("TRANSP:TRANSPARENT")
	[void]$sb.AppendLine(‘END:VEVENT’)
    echo $x.Subject
    echo 'Appointment event saved.'
}
#Once we’ve defined our event, we close out the “objects”.	
[void]$sb.AppendLine(‘END:VCALENDAR’)
echo 'Saving Appointment to .ics file!'

$fileName = 'My_Outlook_Calendar.ics'
$sb.ToString() | Out-File $fileName
echo 'Successful Outlook Calendar Export! Please upload My_Outlook_Calendar.ics to Outlook 365 using \"Add Calendar\" wizard.'
