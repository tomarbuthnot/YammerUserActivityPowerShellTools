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
    Version							: 0.2
 
    Author/Copyright		: Copyright Tom Arbuthnot - All Rights Reserved
    
    Email/Blog/Twitter	: tom@tomarbuthnot.com lyncdup.com @tomarbuthnot
    
    Dedicated Post			: 
    
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
    [Parameter(Mandatory=$false,
    HelpMessage='Users.csv From Yammer Web Export https://about.yammer.com/success/activate/admin-guide/managing-your-users/export-users/')]
    $UsersCSV = 'defaultvalue1',
    
    
    [Parameter(Mandatory=$false,
    HelpMessage='Messages.csv from Yammer Data Web Export https://about.yammer.com/success/activate/admin-guide/monitoring-your-data/export-data/')]
    $MessagesCSV = 'defaultvalue1',
    
    [Parameter(Mandatory=$false,
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
    
    # Script Level Variable to Stop Execution if there is an issue with any stage of the script
    $script:EverythingOK = $true
    
    $Data = Import-Csv "$data"
    $Data | Sort-Object sender_Name | Select-Object Sender_Name,created_at
    
    Write-Debug "Data Variable contains"
    Write-Debug "$Data"

    $UniqueUsers = (Import-Csv $UsersExport) | Where-Object {$_.state -eq 'active'} | Select-Object name,email
    
    $UsersthatPosted = $Data | Select-Object sender_Name -Unique
    
    # Create New Output collection
    $OutputCollection=  @()
    
    
    #############################################################
    # Function to Deal with Error Output to Log file
    #############################################################
    
    Function ErrorCatch-Action 
    {
      Param 	(
        [Parameter(Mandatory=$false,
        HelpMessage='Switch to Allow Errors to be Caught without setting EverythingOK to False, stopping other aspects of the script running')]
        # By default any errors caught will set $EverythingOK to false causing other parts of the script to be skipped
        [switch]$SetEverythingOKVariabletoTrue
      ) # Close Parameters
      
      # Set EverythingOK to false to avoid running dependant actions
      If ($SetEverythingOKVariabletoTrue) {$script:EverythingOK = $true}
      else {$script:EverythingOK = $false}
      Write-Verbose "EverythingOK set to $script:EverythingOK"
      
      # Write Errors to Screen
      Write-Error $Error[0]
      # If Error Logging is runnning write to Error Log
      
      if ($LogErrors) {
        # Add Date to Error Log File
        Get-Date -format 'dd/MM/yyyy HH:mm' | Out-File $ErrorLog -Append
        $Error | Out-File $ErrorLog -Append
        '## LINE BREAK BETWEEN ERRORS ##' | Out-File $ErrorLog -Append
        Write-Warning "Errors Logged to $ErrorLog"
        # Clear Error Log Variable
        $Error.Clear()
      } #Close If
    } # Close Error-CatchActons Function
    
  } #Close Function Begin Block
  
  #############################################################
  # Process Block
  #############################################################
  
  Process {
    
    # First Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        Foreach ($Post in $Data)
        {
          # Generate Real Date
          $StringDate = $Post.created_at | Select-String -Pattern '^(((?<1>[0-9]{4}[/.-](?:1[0-2]|0[1-9])[/.-](?:3[01]|[12][0-9]|0[1-9]))))'
          [datetime]$PostDate = $StringDate.Matches[0].Value
          
          $output = New-Object -TypeName PSobject 
          $output | add-member NoteProperty 'Sender_Name' -value $($Post.Sender_Name)
          $output | add-member NoteProperty 'sender_email' -value $($Post.sender_email)
          $output | add-member NoteProperty 'Date_Posted' -value $($PostDate)
          $output | add-member NoteProperty 'api_url' -value $($Post.api_url)
          $OutputCollection += $output     
        }
        
        # Create Section output collection
        $OutputCollection2=  @()
        
        
        Foreach ($User in $UniqueUsers)
        {
          # Build Date since Last Post
         
          # Clear output
          $LastPostPerUser = $null
          
          $LastPostPerUser  = $OutputCollection | Where-Object {$_.Sender_Name -eq "$($user.Name)"} | Sort-Object Date_Posted -Descending | Select-Object -First 1
          
          If ($LastPostPerUser -ne $null)
          {
            $DateSinceLastPost = $(Get-Date) - $($LastPostPerUser.Date_Posted)
            
            #Build Web URL
            $APIURL = $LastPostPerUser.api_url
            $WebURL = $APIURL.Replace('api/v1',"$NetworkDomain")
            
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
          If ($LastPostPerUser -eq $null)
          {
            Write-Host "No Post for $($user.Sender_Name) $($user.email) in Range Downloaded"
            
          }
          
        }
        
        # Write Output collection to pipeline
        $OutputCollection2
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If Everthing OK Block
    
    #############################################################
    # Next Script Action or Try,Catch Block
    #############################################################
    
    # Second Code To Run
    If ($script:EverythingOK)
    {
      Try 	
      {
        
        # Code Goes here
        
        
      } # Close Try Block
      
      Catch 	{ErrorCatch-Action} # Close Catch Block
      
      
    } # Close If Everthing OK Block
    
    
  } #Close Function Process Block
  
  #############################################################
  # End Block
  #############################################################
  
  End 	{
    Write-Verbose "Ending $($myinvocation.mycommand)"
  } #Close Function End Block
  
} #End Function






 