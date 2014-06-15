# Yammer-UserActivityFromExport Launcher

# This Fuction is written to be a True Function with no user editing
# This means all user inputs are Paramters
# Some people prefer to permentiny have Parameters set and output Pipes setup
# They can be set in this Launcher File


# Variables

$UsersCSV = 'C:\temp\YammerExport\Users.csv'

$MessagesCSV = 'C:\temp\YammerExport\Messages.csv'

$NetworkDomain = 'modalitysystems.com'

# Get the Directory of the Script, note this will not work with Invoke-Command
# If you need Invoke-Command use the Function rather than the laucher
$scriptDir = Split-Path -Parent $myinvocation.mycommand.path

# Load Function
. "$scriptDir\Yammer-UserActivityFromExport.ps1"

# Paramters and Pipe

$output = Yammer-UserActivityFromExport -UsersCSV $UsersCSV -MessagesCSV $MessagesCSV -NetworkDomain $NetworkDomain

$output | Select-Object Sender_Name,Days_Since_Last_Post,sender_email | Sort-Object Days_Since_Last_Post -Descending | ft -AutoSize

Write-Verbose "Emails of users who have not posted in lat 14 days"

$output | Where-Object {$_.Days_Since_Last_Post -gt 14} | Select-Object sender_email