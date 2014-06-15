Yammer-UserActivityFromExport
=============================

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
 
    Author/Copyright	: Copyright Tom Arbuthnot - All Rights Reserved
    
    Email/Blog/Twitter	: tom@tomarbuthnot.com lyncdup.com @tomarbuthnot
    
	Dedicated Post	: https://github.com/tomarbuthnot/Yammer-UserActivityFromExport
    
    Disclaimer   	: THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
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
    
     
    Acknowledgements 	: 
   
    Assumptions		: ExecutionPolicy of AllSigned (recommended), RemoteSigned or Unrestricted (not recommended)
    
    Limitations		:    										
    
    Known issues	: 

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
