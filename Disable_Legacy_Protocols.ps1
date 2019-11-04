<#
.SYNOPSIS
Set-DelPermSubfolders.ps1 grants certain permission set to a mailbox folder's and subfolders
as defined in an input file.

.DESCRIPTION 
The script reads mailbox information from a CSV input file, finds a single mailbox 
(matching Identity column), and grants the user (as specified in the GrantToUser column)
certain permissions (as specified in the PermissionSet).

.PARAMETER ConfigFile
Input CSV file that must contain the following information:

 * UserPrincipalName - mailbox UPN such as primary SMTP address, account, etc.
 * GrantToUser - user account that should be granted permissions,mailbox UPN such as primary SMTP address, account, etc.
 * AccessRightsFolder - permissions to apply separated by semicolon if needed


.PARAMETER FolderPermission
Gathering all the mailboxe folder which are assign with permissions and export to a csv file 

.PARAMETER OutputDir
Output directory for storing output and log files. If not specified output is 
generated in '.\SubFolders'. A separate sub-directory is created for OutputFiles 
and Logs. 

.EXAMPLE
.\Set-DelPermSubfolders.ps1

.NOTES
AUTHOR
Joanna Vathis, jovath@microsoft.com

COPYRIGHT
(c) 2017 Microsoft, all rights reserved

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

CHANGE LOG
v1.0, 2017-03-23 - Initial version
v1.1, 2017-04-24 - Bug fixing
v2.0, 2017-08-02 - Correct the values in the script
#>


Param (
[string]$Results = [string]::Format(".\IMAP_POP_Output_{0}.csv", (Get-Date).Tostring("yyyy-MM-dd_HHmmss")))


#===============================================================================
# Function: LoginOffice365
# Connect to Office 365 and Exchange Online
#===============================================================================

 Function Create-PSSession (){

Clear
Do {
    Do {
  
  # Check if Pssesion Open and if its open, remove it
    Get-PSSession | Remove-PSSession
 
 # Check if the O365 Username is not empty as viriable 
  If($O365Username -eq $Null -or $ConvertOffice365Password -eq $Null) {
  
  # Write a message
    Write-Host @"
.......Before we get started, please enter your Office 365 Administrative username & password......
"@ -ForegroundColor Blue -BackgroundColor White
   
 # Set variables
    $O365Username = Read-Host "Enter your Office 365 Username"
    $O365Password = Read-Host -AsSecureString "Enter your Office 365 Password"
    $ConvertOffice365Password = ConvertFrom-SecureString($O365Password)
 
 }else{ 

    $CloudPassword = ConvertTo-SecureString($ConvertOffice365Password)        
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $O365Username, $O365Password
    $Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    Connect-MsolService –Credential $cred
    Write-Host " "
    Write-Host " "
    Write-Host " "

    }
}
   While ($Cred -eq $Null)
   $FindPSSession = Get-PSSession
 } While ($FindPSSession -eq $Null)

}


# ===============================================================================================================================================================
#                                                             SCRIPT BODY - MAIN CODE                                                                             
# ===============================================================================================================================================================

# Set Variables
  $logname = Get-Date -Format "MM.dd.yyyy_HH.mm"
  $logpath = ".\O365Logs_$logname.log"


# Start Transcript
  Start-Transcript -Path $logpath -Append

# Create PSSession to Exchange
   $session = Get-PSSession -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    If(($session -ne $null) -and ($session.State -eq 'Opened') -and ($session.Availability -eq 'Available') -and
       ($session.ConfigurationName -eq 'Microsoft.Exchange') -and ($session.ComputerName -eq "outlook.office365.com")){
        Write-Host "Reusing opened PSSession to  outlook.office365.com" -ForegroundColor Cyan
        Write-Host "*******************************************************" -ForegroundColor White
        Write-Host " "
  

} elseif (($session -eq $null) -or ($session.State -eq 'Broken')){

# Create a new Session to Office 365
  Write-Host "Create a new PSSession to Exchange Online......." -ForegroundColor Yellow
  Write-Host "*****************************************************" -ForegroundColor White
  Write-Host " "
  Create-PSSession -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null     
}

[Array]$mbx = @()

# Collect all Mailboxes 
  $mbx = Get-Mailbox -ResultSize Unlimited

# Loop and disable IMAP and POP protocol
   Foreach ($m in $mbx){
    Get-CASMailbox | Where-Object {($_.PopEnabled -eq $true) -and ($_.ImapEnabled -eq $true)}
    Set-CASMailbox -Identity $m.alias -PopEnabled:$false -ImapEnabled:$false 
 }

# Collect and export the results after disable the IMAP and POP protocols
$output = Get-CASMailbox | Where-Object {($_.PopEnabled -eq $false)} 
$output |select Name,PopEnabled,ImapEnabled | Export-Csv -Path $Results -NoTypeInformation -Encoding UTF8
Stop-Transcript -ErrorAction SilentlyContinue -WarningAction SilentlyContinue