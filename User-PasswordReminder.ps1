<#
.SYNOPSIS
	Send email reminders to users with expiring passwords and create a CSV and HTML Report for admins.

.DESCRIPTION
	This script will send email reminder to users whose passwords are expiring
    within the specified time or have already expired. You also have the option
    of sending an email to your helpdesk if the days until expiration drops below
    a specific level.
    
    This script will also create a report in two formats; CSV and HTML. Both reports 
    are then emailed to the specified user.

.PARAMETER Path
	Specify the path where you want to save the CSV report.  The script does not save
	the HTML report, but emails it as the body of the email.

.PARAMTER SearchBase
	Specify the full FQDN of the OU you wish to begin your search at.  If you need
	help getting the FQDN you can use this script:
	http://community.spiceworks.com/scripts/show/1635-copy-a-ou-s-fqdn-to-clipboard

.PARAMETER RemindAge
	Specify the number of days the script will report on expiring passwords.  An
	entry of 2 will report all passwords that have already expired, or are going to
	expire in the next 2 days.

.PARAMETER TicketAge
	Specify the number of days the script will report on expiring passwords.  An
	entry of 2 will report all passwords that have already expired, or are going to
	expire in the next 2 days.

.PARAMETER From
	When emailing the reports, this parameter will specify who the report is 
	coming from. 
 
.PARAMETER AdminEmail
	Tell the script which email address to use for testing

.PARAMETER ReportEmail
	Tell the script where to email the reports that are generated

.PARAMETER SMTPServer
	This needs to be the IP address or name of your SMTP relay server.

.PARAMETER TestMode
	This setting will enable test much which will send all emails to AdminEmail and display results to console

.PARAMETER OpenTicket
	Tell the script where to send emails to your help desk

.PARAMETER NoReport
	Tell the script not to generate or send reports

.PARAMETER ReportOnly
	Tell the script to only generate and send reports 

.PARAMETER DisplayResults
	Tell the script to display the report data to the console

.OUTPUTS
	CSV:	ExpirationReport.csv in the $Path location
	Email:	HTML version of the same report in the body of the email.
            It also attaches the CSV to the email.

.NOTES
	Script:				User-PasswordReminder.ps1
    Version:            1.0
	Author:				Derrick Lamers
	Creation Date:		1/17/2018
    Website:            www.lamerstech.com

.EXAMPLE
	.\User-PasswordReminder.ps1
	Accepts all defaults as defined in the PARAM section

.EXAMPLE
	.\Report-PasswordReminder.ps1 -Path C:\Reports -RemindAge 15 -From support@yourdomain.com -ReportEmail tech@yourdomain.com -SMTPServer smtp.onmicrosoft.com
	Runs the report using C:\Reports as the path to save the CSV report.  Email will be sent
	from "support@yourdomain.com" and sent to "tech@yourdomain.com" using smtp.onmicrosoft.com
	as the SMTP relay server.  All user accounts with expired passwords or password that will
	be expiring in 15 days will be reported and sent an email reminder.
#>
Param (
	[string]$Path = "C:\TEMP",
	[string]$SearchBase = "DC=net,DC=example,DC=com",
    [int]$RemindAge = 15,
	[int]$TicketAge = 8,
	[string]$From = "",
	[string]$AdminEmail = "",
    [string]$ReportEmail = "",
    [string]$TicketEmail = "",
	[string]$SMTPServer = "",
    [switch]$TestMode,
    [switch]$OpenTicket,
    [switch]$NoReport,
    [switch]$ReportOnly,
    [switch]$DisplayResults
)

#Array that will store all account records
$Result = @()

#Get today's date
$CurrentDate = [datetime]::Now.Date

#The message body used when creating help desk tickets
$OpenTicketMessage = @"
@@CATEGORY=User Administration@@
@@SUBCATEGORY=Network@@
@@ITEM=Password Expiration@@
@@LEVEL=Tier 1@@
@@MODE=Automated@@
@@PRIORITY=Low@@
"@

cls

function GetDaysToExpire([datetime] $expireDate)
{
	# Function: Expects the remaining Days to Expire
    $date =  New-TimeSpan $CurrentDate $expireDate
    return $date.Days
}

function EmailUser
{
	# Function: Send Email to User
    PARAM ($User)
    $today = $CurrentDate.ToString("dddd (MM/dd/yyyy)")
    $daystoexpire = GetDaysToExpire $User.ExpirationDate
	$SmtpClient = New-Object system.net.mail.smtpClient 
 	$mailmessage = New-Object system.net.mail.mailmessage 
 	$SmtpClient.Host = $SMTPServer 
 	$mailmessage.From = $From
    #If Test Mode is activated, only send emails to AdminEmail
    If ($TestMode -eq $false)
    {
        $mailmessage.To.add($AdminEmail)
        If ($RemindAdmin)
        {
           $mailmessage.To.add($AdminEmail)
        }
    } 
    else
    {
      $mailmessage.To.add($AdminEmail)
    }  

    $mailmessage.IsBodyHtml = $true

    If( $daysToExpire -le 0 )
    {
 	    # When the Days to Expire is zero, this email header will be sent
	    $mailmessage.Subject = "Your password expired on " + $User.ExpirationDate
	    $mailmessage.Body += "Hi " + $User.Firstname + ","
	    $mailmessage.Body += " <br /><br />"
	    $mailmessage.Body += "This email is to inform you that your network password will expired on <font color=red>" + $User.ExpirationDate + " at " + $User.ExpirationTime + "</font>. Access to your computer, email, and other NBS resources will be affected until your password is changed."
    }
    else
    {
        # When the Days to Expire is more than zero, this email header will be sent
	    $mailmessage.Subject = "Your NBS password expires on " + $User.ExpirationDate
	    $mailmessage.Body += "Hi " + $User.Firstname + ","
	    $mailmessage.Body += "<br /><br />"
	    $mailmessage.Body += "This email is to reminder you that your network password will expire on <font color=red>" + $User.ExpirationDate + " at " + $User.ExpirationTime +"</font>. Access to your computer, email, and other NBS resources will be affected if your password is not changed prior to this date."
    }

    # Body of the email sent to users
	$mailmessage.Body += "<br /><br />"
	$mailmessage.Body += "If you don't have access to an computer to change your password, please contact <a href='mailto:itsupport@yournbs.com?subject=Need Help Changing My Password'>IT Support</a> and we will be happy to assist you in changing your password."
	$mailmessage.Body += "<br /><br />"	
	$mailmessage.Body += "<strong><u>How do I change my password?</u></strong><br /><br />"
	$mailmessage.Body += "Changing your password is easy and simple but can only be done will working at an office."
	$mailmessage.Body += "<br /><br /><ol>"
	$mailmessage.Body += "<li>Login to your computer as normal (Do not login to RDS)</li>"
	$mailmessage.Body += "<li>Press the <strong>CTRL</strong> + <strong>ALT</strong> + <strong>DELETE</strong> on your keyboard at the same time</li>"
    $mailmessage.Body += "<li>Select Change a Password</li>"
	$mailmessage.Body += "<li>Enter your current password</li>"
	$mailmessage.Body += "<li>Enter your new Password in the <strong>New Password</strong> box</li>"
    $mailmessage.Body += "<li>Enter your new password again in the <strong>Confirm Password</strong> box</li>"
    $mailmessage.Body += "<li>Press <strong>ENTER</strong></li></ol>"
	$mailmessage.Body += "The computer will take a moment to update your new password and then confirm that your password change has been accepted. Once accepted, you may now login to your computer and RDS using your new password."
	$mailmessage.Body += "<br /><br />"
	$mailmessage.Body += "<strong><u>Don't Forget - Update Your Mobile Device(s)</u></strong>"
	$mailmessage.Body += "<br /><ul>"
	$mailmessage.Body += "<li>If you have NBS email configured on your mobile device, you will need to update the password within your devices email client.</li>"
	$mailmessage.Body += "<li>If you connect to the Mobile network with your mobile device, you may be prompted to provide your new password. If you experience issues connecting to Mobile after changing your password, it is recommended that you forget the network on your device and reconnect using your username and new password.</li>"
	$mailmessage.Body += "</ul>"
	$mailmessage.Body += "<strong><u>When changing my password, my computer says it doesn't meet the complexity requirements, what does this mean?</u></strong>"
	$mailmessage.Body += "<br /><br />"
	$mailmessage.Body += "Your new password must meet the following requirements:<br />"
	$mailmessage.Body += "<ul><li>Must be at least 10 characters in length</li>"
	$mailmessage.Body += "<li>Must contain at least one uppercase character (A-Z)</li>"
	$mailmessage.Body += "<li>Must contain at least one lowercase character (A-Z)</li>"
	$mailmessage.Body += "<li>Must contain at least one number character</li>"
    $mailmessage.Body += "<li>Must contain at least one special character (!, $, #, %)</li>"
    $mailmessage.Body += "<li>Must not be the same as your last 5 NBS passwords</li></ul>"
	$mailmessage.Body += "If you have any questions or concerns, please contact <a href='mailto:itsupport@example.com?subject=Need Help Changing My Password'>IT Support</a>."
    $mailmessage.Body += "<br /><br />"
    $mailmessage.Body += "Sincerely,"
    $mailmessage.Body += "<br /><br />"
    $mailmessage.Body += "<strong>The IT Support Team</strong>"
	$mailmessage.Body += "<br /><br />"
	$mailmessage.Body += "------------------------------------------------------------------"
	$mailmessage.Body += "<br />"
	$mailmessage.Body += "Message generated on: " + "<font color=red> $today </font color = red> <br />"

    $smtpclient.Send($mailmessage)

}

function OpenTicket($User)
{
	$SmtpClient = New-Object system.net.mail.smtpClient 
 	$mailmessage = New-Object system.net.mail.mailmessage 
 	$SmtpClient.Host = $SMTPServer 
 	$mailmessage.From = $User.Email 
    $mailmessage.To.add("itsupport@example.com")
    $mailmessage.IsBodyHtml = $true
    $mailmessage.Subject = "@@@Your password expires on " + $User.ExpirationDate
    $mailmessage.Body = $OpenTicketMessage
    $smtpclient.Send($mailmessage)
}

function EmailReport($Result)
{
#Produce a CSV
$ExportDate = Get-Date -f "yyyy-MM-dd"
$Result | Export-Csv $path\AccountExpirationReport-$ExportDate.csv -NoTypeInformation

#Send HTML Email
$Header = "<style>TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;} TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;} TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}</style>"
$splat = @{
	From = $From
	To = $ReportEmail
	SMTPServer = $SMTPServer
	Subject = "Password Expiration Report " + $CurrentDate.ToString("(MM/dd/yyyy)")
}
$Body = $Result | ConvertTo-Html -Head $Header | Out-String
Send-MailMessage @splat -Body $Body -BodyAsHTML -Attachments $Path\AccountExpirationReport-$ExportDate.csv
}

##############################################################################################################################################################################################
######################################################################### MAIN APPLICATION ###################################################################################################
##############################################################################################################################################################################################

#Determine MaxPasswordAge
$maxPasswordAgeTimeSpan = $null
$dfl = (get-addomain).DomainMode
$maxPasswordAgeTimeSpan = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
If ($maxPasswordAgeTimeSpan -eq $null -or $maxPasswordAgeTimeSpan.TotalMilliseconds -eq 0) 
{	Write-Host "MaxPasswordAge is not set for the domain or is set to zero!"
	Write-Host "So no password expiration's possible."
	Exit
}

$Users = Get-ADUser -Filter * -SearchBase $SearchBase -SearchScope Subtree -Properties GivenName,sn,PasswordExpired,PasswordLastSet,PasswordneverExpires,LastLogonDate,Mail
ForEach ($User in $Users)
{	If ($User.PasswordNeverExpires -or $User.PasswordLastSet -eq $null -or $User.Enabled -eq $false)
	{	Continue
	}
	If ($dfl -ge 3) 
	{	## Greater than Windows2008 domain functional level
		$accountFGPP = $null
		$accountFGPP = Get-ADUserResultantPasswordPolicy $User
    	If ($accountFGPP -ne $null) 
		{	$ResultPasswordAgeTimeSpan = $accountFGPP.MaxPasswordAge
    	} 
		Else 
		{	$ResultPasswordAgeTimeSpan = $maxPasswordAgeTimeSpan
    	}
	}
	Else
	{	$ResultPasswordAgeTimeSpan = $maxPasswordAgeTimeSpan
	}
	$Expiration = $User.PasswordLastSet + $ResultPasswordAgeTimeSpan
    $DaysRemaining = (New-TimeSpan -Start (Get-Date) -End $Expiration).Days
    If ($DaysRemaining -le 0)
    {
        $DaysRemaining = 0
    }
	If ($DaysRemaining -le $RemindAge)
	{	$Result += New-Object PSObject -Property @{
			'LastName' = $User.sn
			'FirstName' = $User.GivenName
			UserName = $User.SamAccountName
            'Email' = $User.Mail
			'ExpirationDate' = $Expiration.ToString("d")
            'ExpirationTime' = $Expiration.ToString("t")
            'DaysRemaining' = $DaysRemaining
			'LastLogonDate' = $User.LastLogonDate
			State = If ($User.Enabled) { "Enabled" } Else { "Disabled" }
		}
	}
}
$Result = $Result | Select 'LastName','FirstName',UserName,'Email','ExpirationDate','ExpirationTime','DaysRemaining','LastLogonDate',State | Sort 'ExpirationDate','ExpirationDate','LastName'

If ($TestMode -or $DisplayResults)
{
    $Result | Format-Table -AutoSize
}

If ($ReportOnly -eq $false)
{
    foreach( $User in $Result ) 
    { 
        If ($User.Email -ne $null)
        {
            If ($User.DaysRemaining -gt $TicketAge)
            {
                #Write-Host "Email User"
                #EmailUser $User
            }
            else
            {
                if ($OpenTicket -eq $false)
                {
                    #Write-Host $User.Email" | Open Ticket"
                    #OpenTicket $User
                }
            }
        }
    }
}

#Generate and email reports
if ($NoReport = $false)
{
    EmailReport $Result
}
