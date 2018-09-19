##################################
#Email Report Script
#Runs Once per day as scheduled Task
#Requires exchange management console
#Gives snapshot of high volume mail
#################################


#### Exchange connection Chunk  ####
Write-Verbose "Loading the Exchange snapin"

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

. $env:ExchangeInstallPath\bin\RemoteExchange.ps1

Connect-ExchangeServer -auto -AllowClobber
##############################

$sent_mail_threshold = 150
$domain_threshold = 20
$whitelist_path = 'C:\Users\rshannon3\Desktop\Email_Reports\email_whitelist.txt'
$whitelist_contents = @()
$start_date=[DateTime]::Now.AddHours(-30).toString()
$end_date=[DateTime]::Now.AddHours(-6).toString()


foreach($line in get-content $whitelist_path) { $whitelist_contents += $line }

#retrive all mail in last 24 hours
$global:recent_mail = Get-TransportServer | Get-MessageTrackingLog -Start $start_date  -end $end_date -EventId "Send" -resultsize unlimited |
    select timestamp, sender, recipients, messagesubject

#sort by count above threshold and not in whitelist
$global:recent_senders = $recent_mail | Group-Object sender | where {$_.name -like "*.redacted*" -and -not ( $whitelist_contents -contains $_.name -or $_.name -like 'Microsoft*@redacted*' ) } | sort count -Descending


#Report table
$table = @()

#Populate Table
for($c = 0; $c -lt $recent_senders.Count; $c++)
{
    $count_column = 0 #$recent_senders.get($c).count
    $name_column = $recent_senders.get($c).name

    $subjects_hash=@{}

    $top_subjects=$recent_senders.get($c).group | group messagesubject <#where {$_.count -gt $sent_mail_threshold}#> | sort count -Descending #| Select-Object -First 1
    #Add threshold counter to include num of recipients
    $adjusted_message_count = 0;
    $top_subjects | ForEach-Object {
        $adjusted_message_count += $_.group.recipients.Count
        $subjects_hash.add($_.Name, $_.Group.recipients.Count)
    }
    $count_column = $adjusted_message_count


    $subject_column = ($subjects_hash.GetEnumerator() | Sort-Object -Property value -Descending | where {$_.Value -gt $sent_mail_threshold} | % { "$($_.Key) = $($_.Value)" }) -join "`n"

    if( $adjusted_message_count -lt $sent_mail_threshold -or $subject_column.Length -lt 1) #If no subject sent over threshold , do not add this sender to the report
    { continue }

    $domains=@{}  ## Summary of each recipient domain. Stored in a hash format      domain : number of emails

    $recent_senders.group | where {$_.sender -eq $recent_senders.get($c).name} | #Populate table
        foreach-object {
            $_.recipients |
                foreach-object {
                    $ind = $_.IndexOf("@")
                    $dom = $_.Substring($ind+1)
                    if($domains.ContainsKey($dom)) {
                        $domains.Set_Item($dom, $domains.Get_Item($dom) + 1 )
                        }
                    else {
                        $domains.Add($dom, 1)
                    }
                }
        }

    $domain_column = ($domains.GetEnumerator() | Where-Object {$_.Value -gt $domain_threshold} | Sort-Object -Property value -Descending | % { "$($_.Key) = $($_.Value)" }) -join "`n"    #Sort and convert hashtable to string
    $table += ,@($count_column, $name_column, $subject_column, $domain_column) #Store in table

}
$table = $table | Sort-Object @{Expression={$_[0]}} -Descending
$path = "C:\Users\rshannon3\Desktop\Email_Reports\email_report $(get-date -format MM-dd-yyyy).txt"

Out-File $path -InputObject "Date From: $start_date To: $end_date              `n Current Subject Threshold: $sent_mail_threshold `n" #Date Header
if( $table.Count -lt 1) {
    Out-File $path -InputObject "No entries found. `n" -Append
}
else {
    Format-Table -Wrap -AutoSize -InputObject $table @{label="Count"; expression={$_[0]}},@{label="Sender"; expression={$_[1]}},@{label="Top Subject"; expression={$_[2]}},@{label="Domain Summary"; expression={$_[3]}} |
        Out-File $path -Append
}

#############################################
$username = ""
$password = ""
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

$From = ""
$To = ""
$self = ""
$Attachment = "C:\Users\rshannon3\Desktop\Email_Reports\email_report $(get-date -format MM-dd-yyyy).txt"
$Subject = "Email Report $(get-date -format MM-dd-yyyy)"
$Body = "Report is attached."
$SMTPServer = "smtp.gmail.com"
$SMTPPort = "587"
Send-MailMessage -From $From -to $To,$self -Subject $Subject `
-Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
-Credential $cred -Attachments $Attachment

###########################################################>
