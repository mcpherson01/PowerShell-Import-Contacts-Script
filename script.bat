
: ---> IMPORT OUTLOOK CONTACTS
powershell -Command DRIVELETTER:\LOCATION\ImportContacts.ps1

:----> Make a seperate ps1 

$olFolderCalendar = 9
$olAppointmentItem = 1
$Outlook = New-Object -comobject Outlook.Application
$objNamespace = $Outlook.GetNamespace("MAPI")
$objCalendar = $objNamespace.GetDefaultFolder($olFolderCalendar)
$projectlist = Import-CSV "DRIVELETTER:\LOCATION\CONTACTS.csv"

$q = $projectlist.count
for ($i = 0; $i -le $q; $i++) {
	$h = $outlook.CreateItem($olAppointmentItem)
	$h.subject = $projectlist[$i].subject
	$h.start = $projectlist[$i].'start date'
	$h.end = $projectlist[$i].'end date'
	$h.location = $projectlist[$i].location 
	$h.AllDayEvent = $True
	$h.ReminderSet = $False
	$a = $h.save()
} 

