# Yammer-UserActivityFromExport Launcher

# This Fuction is written to be a True Function with no user editing
# This means all user inputs are Paramters
# Some people prefer to permentiny have Parameters set and output Pipes setup
# They can be set in this Launcher File


# Variables

$UsersExport = 'C:\temp\YammerExport\Users.csv'

$Data = 'C:\temp\YammerExport\Messages.csv'

$NetworkDomain = 'modalitysystems.com'

# Get the Directory of the Script, note this will not work with Invoke-Command
# If you need Invoke-Command use the Function rather than the laucher
$scriptDir = Split-Path -Parent $myinvocation.mycommand.path

# Load Function
. "$scriptDir\Yammer-UserActivityFromExport.ps1"

# Paramters and Pipe

# Grab output into an object
# Yammer-UserActivityFromExport | Select-Object Sender_Name,Days_Since_Last_Post,sender_email | Sort-Object Days_Since_Last_Post -Descending | ft -AutoSize

# $output = Yammer-UserActivityFromExport

Yammer-UserActivityFromExport -Debug