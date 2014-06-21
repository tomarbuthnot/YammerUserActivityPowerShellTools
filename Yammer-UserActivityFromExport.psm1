Function Yammer-UserActivityFromExport {
<#
.SYNOPSIS  
		Takes Yammer CSV Exports for Users and Messages and Parses them to get the last post date and URL for Each User
    Allow those less active to be chases for updates/spot training opportunities

.DESCRIPTION  
		Takes Yammer CSV Exports for Users and Messages and Parses them to get the last post date and URL for Each User
    Allow those less active to be chases for updates/spot training opportunities

    Yammer Export Users
    https://about.yammer.com/success/activate/admin-guide/managing-your-users/export-users/


.NOTES  
    Version							: 0.3
 
    Author/Copyright		: Copyright Tom Arbuthnot - All Rights Reserved
    
    Email/Blog/Twitter	: tom@tomarbuthnot.com lyncdup.com @tomarbuthnot
    
    Dedicated Post			: https://github.com/tomarbuthnot/Yammer-UserActivityFromExport
    
    Disclaimer   				: THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
                          OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
                          While these scripts are tested and working in my environment, it is recommended 
                          that you test these scripts in a test environment before using in your production 
                          environment
                          Tom Arbuthnot further disclaims all implied warranties including, without limitation, any 
                          implied warranties of merchantability or of fitness for a particular purpose. The entire risk 
                          arising out of the use or performance of this script and documentation remains with you. 
                          In no event shall Tom Arbuthnot, its authors, or anyone else involved in the creation, production, 
                          or delivery of this script/tool be liable for any damages whatsoever (including, without limitation, 
                          damages for loss of business profits, business interruption, loss of business information, or other 
                          pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, 
                          even if Tom Arbuthnot has been advised of the possibility of such damages.
    
     
    Acknowledgements 		: 
    
    Assumptions					: ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
    Limitations					:    										
    
    Known issues				: 

    Ideas/Wish list			: # Use Rest API to pull the exports based on date range
    

.LINK  
    http://link.com

# Examples

.EXAMPLE
		Yammer-UserActivityFromExport
 
		Description
		-----------
		Returns Objects

.EXAMPLE
		Yammer-UserActivityFromExport | Select-Object Sender_Name,Days_Since_Last_Post,sender_email | Sort-Object Days_Since_Last_Post -Descending | ft -AutoSize
 
		Description
		-----------
		Returns All Users sorted by Date Last Posted

#>
  
  
  #############################################################
  # Param Block
  #############################################################
  
  # Sets that -Whatif and -Confirm should be allowed
  [cmdletbinding(SupportsShouldProcess=$true)]
  
  Param 	(
    [Parameter(Mandatory=$true,
    HelpMessage='Users.csv From Yammer Web Export https://about.yammer.com/success/activate/admin-guide/managing-your-users/export-users/')]
    $UsersCSV = 'defaultvalue1',
    
    [Parameter(Mandatory=$true,
    HelpMessage='Messages.csv from Yammer Data Web Export https://about.yammer.com/success/activate/admin-guide/monitoring-your-data/export-data/')]
    $MessagesCSV = 'defaultvalue1',
    
    [Parameter(Mandatory=$true,
    HelpMessage='domain.com , used in a find/replace to build clickable URLs to posts rather than API URLs')]
    $NetworkDomain = 'defaultvalue1',
    
    [Parameter(Mandatory=$false,
    HelpMessage='Error Log location, default C:\<Command Name>_ErrorLog.txt')]
    [string]$ErrorLog = "c:\$($myinvocation.mycommand)_ErrorLog.txt",
    [switch]$LogErrors
    
  ) #Close Parameters
  
  
  #############################################################
  # Begin Block
  #############################################################
  
  Begin 	{
    Write-Verbose "Starting $($myinvocation.mycommand)"
    Write-Verbose "Error log will be $ErrorLog"
    
    # Parameters input
    Write-Verbose "Messages CSV to import is $MessagesCSV"
    Write-Verbose "UserCSV is $UsersCSV"
    Write-Verbose "Networkdomain is $NetworkDomain"

    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    Write-Verbose "UserCSV is $UsersCSV"

    $MessagesImport = Import-Csv "$MessagesCSV"
    
    $Data = $MessagesImport  | Sort-Object sender_Name | Select-Object Sender_Name,created_at,api_url,sender_email

    $UniqueUsers = (Import-Csv $UsersCSV) | Where-Object {$_.state -eq 'active'} | Select-Object name,email
    
    Write-Verbose "Unique Users Count $($UniqueUsers.count)"

    $UsersthatPosted = $Data | Select-Object sender_Name -Unique
    
    # Create New Output collection
    $OutputCollection=  @()

    
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK -eq $True)
    {
      Try 	
      {
        
        # Create an Object with Each Post with datetimeobject with posters details
        Foreach ($Post in $Data)
        {
          # Generate Real Date
          $StringDate = $Post.created_at | Select-String -Pattern '^(((?<1>[0-9]{4}[/.-](?:1[0-2]|0[1-9])[/.-](?:3[01]|[12][0-9]|0[1-9]))))'
          [datetime]$PostDate = $StringDate.Matches[0].Value
          
          # Write-Verbose "Date of Post was $PostDate, user posting was $($Post.Sender_Name)"

          $output = New-Object -TypeName PSobject 
          $output | add-member NoteProperty 'Sender_Name' -value $($Post.Sender_Name)
          $output | add-member NoteProperty 'sender_email' -value $($Post.sender_email)
          $output | add-member NoteProperty 'Date_Posted' -value $($PostDate)
          $output | add-member NoteProperty 'api_url' -value $($Post.api_url)
          $OutputCollection += $output     
        }
        
        # Create Section output collection
        $OutputCollection2=  @()
        

        # Create an Object with each users last post
        Foreach ($User in $UniqueUsers)
        {
          Write-Verbose "Working on $($user.Name)"
          # Build Date since Last Post
         
          # Clear output
          $LastPostPerUser = $null
          
          $LastPostPerUser  = $OutputCollection | Where-Object {$_.Sender_Name -eq "$($user.Name)"} | Sort-Object Date_Posted -Descending | Select-Object -First 1 -ErrorAction SilentlyContinue

          Write-Verbose "Last Post for $user is $LastPostPerUser"

           If ($LastPostPerUser -eq $null)
          {
            Write-Verbose "No Post for $($user.email) in CSV Loaded, will be listed as 999 days"

            $output = New-Object -TypeName PSobject 
            
            $output | add-member NoteProperty 'Sender_Name' -value $($User.Sender_Name)
            $output | add-member NoteProperty 'sender_email' -value $($User.email)
            $output | add-member NoteProperty 'Date_Posted' -value $null
            $output | add-member NoteProperty 'Days_Since_Last_Post' -value '999'
            $output | add-member NoteProperty 'Web_URL' -value $null
            
            # Add output to output collection
            $OutputCollection2 += $output

            
          }
          If ($LastPostPerUser -ne $null)
          {
            $DateSinceLastPost = $(Get-Date) - $($LastPostPerUser.Date_Posted)
            
            #Build Web URL         
            $APIURL = $LastPostPerUser.api_url
            Write-Verbose "API URL is $APIURL"
            $WebURL = $APIURL.Replace('api/v1',"$($NetworkDomain)")
            Write-Verbose "WebURL is $WebURL"

            # create new output object
            $output = New-Object -TypeName PSobject 
            
            $output | add-member NoteProperty 'Sender_Name' -value $($LastPostPerUser.Sender_Name)
            $output | add-member NoteProperty 'sender_email' -value $($LastPostPerUser.sender_email)
            $output | add-member NoteProperty 'Date_Posted' -value $($LastPostPerUser.Date_Posted)
            $output | add-member NoteProperty 'Days_Since_Last_Post' -value $($DateSinceLastPost.Days)
            $output | add-member NoteProperty 'Web_URL' -value $($WebURL)
            
            # Add output to output collection
            $OutputCollection2 += $output
          }
         
          
        }
        
        # Write Output collection to pipeline
        $OutputCollection2
        
        
      } # Close Try Block
      
      Catch {  Write-Verbose "Error hit"
              $Error[0]
            } # Close Catch Block
      
      
    } # Close If EverythingOK Block
    
  } # Close Process Block
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function






 