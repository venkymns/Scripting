Function Write-Log([string]$message)
{
   Out-File -InputObject $message -FilePath $LogFile -Append
}

Function Send-RevokeMail ([Switch]$Extension, [Switch]$Revoked)
{
	$bcc = "grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com"
	$sub = "$REQ`: Access Expiry Alert!!!"
	$data = Get-ADuser -Filter "Samaccountname -eq '$Membername'"	
	if (($mail) -and ($data))
	{
		$from = (($mail -split '@')[0]).Split('.')[0]
		$fn = $data.GivenName
		$sn = $data.Surname
		[string[]]$to = "$fn.$sn@marks-and-spencer.com","$mail"
		$Body = "Hi <b>$fn</b>/<b>$from</b> <br><br> "
	} elseif (!($mail) -or !($data))
	{
		if ($mail)
		{
			$from = (($mail -split '@')[0]).Split('.')[0]
			$to = $mail
			$Body = "Hi <b>$from</b> <br><br> "
		} elseif ($data)
		{
			$fn = $data.GivenName
			$sn = $data.Surname
			$to = "$fn.$sn@marks-and-spencer.com"
			$Body = "Hi <b>$fn</b> <br><br> "
		}
	} else {
		$to = $bcc
	}
	If ($Extension)
	{
		$body += "Access provided to the request <b>$REQ</b> will be expired on <b>$ExpDate</b> <br> "
		$body += "Please raise an extension request in Non-Prod Remedy portal to avoid access revocation <br><br>"
		$body += "<u><b>Details:</b></u> <br>"
		if ($SP -eq "Sharedpath,MNSUKCATE")
		{
			$body += "<b>Category:</b> $Mailip <br>"
			$body += "<b>Path:</b> $Accpath <br>"
		} else {
			$body += "<b>Category:</b> $Mailip <br>"
			$body += "<b>Group Name:</b> $Groupname <br>"
		}
		$body += "<b>Member Name:</b> $Membername <br>"
		$body += "<b>Domain:</b> $domain <br><br>"
		$body += "Regards,<br>Non-Prod Wintel<br><br>"
		$Body += "*** Please do not reply, This is a system generated email *** <br>"
		$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"		
	} elseif ($Revoked)
	{
		$body += "Access provided to the request <b>$REQ</b> is now expired. If required, please raise a new request in Non-Prod Remedy portal. <br><br>"
		$body += "<u><b>Details:</b></u> <br>"
		if ($SP -eq "Sharedpath,MNSUKCATE")
		{
			$body += "<b>Category:</b> $Mailip <br>"
			$body += "<b>Path:</b> $Accpath <br>"
		} else {
			$body += "<b>Category:</b> $Mailip <br>"
			$body += "<b>Group Name:</b> $Groupname <br>"
		}
		$body += "<b>Member Name:</b> $Membername <br>"
		$body += "<b>Domain:</b> $domain <br><br>"
		$body += "Regards,<br>Non-Prod Wintel<br><br>"
		$Body += "*** Please do not reply, This is a system generated email *** <br>"
		$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"	
	}
	Send-MailMessage -From "Non-Prod Access Management <Non-ProdAccessManagement@marksandspencercate.com>" -To $to -BCC $bcc -Subject $sub -Body $Body -BodyAsHtml -Priority High -SmtpServer 10.151.209.17 -EA Stop
}

Function Start-Main
{
	try
	{
		Import-Module ActiveDirectory -ErrorAction Stop
		$script:Today = "{0:dd-MMM-yy}" -f (get-date)
		#$script:MY = "{0:MMMyy}" -f (get-date)
		#$script:foldate = (Get-date).AddDays(27).day
		#$script:foldate = (Get-date).day
		$script:DBlog = "AccessManagement"
		$script:Currdom = [Environment]::UserDomainName
		#$Script:d = $Domain.Tostring().ToLower()
		$Script:d = $domain.Substring('5')
        $Script:Dcaps = $Domain.Tostring().ToUpper()
		$Script:serverdom = "mnsuk$d.adroot$d.marksandspencer$d.com"
		$Script:DC = "DC=mnsuk$d,DC=adroot$d,DC=marksandspencer$d,DC=com"
		$encrypted = "76492d1116743f0423413b16050a5345MgB8AFIAZAB1AGoARQB4AHQAWgA3AHMAMABVAG4AawBGAFQAKwBQAE4AeQBSAEEAPQA9AHwAYQBmADcAZgBhADYAMgA0ADYAYwAzADMAOAA1AGQAMAA3AGUAMAA5AGMAMABlAGEANAAzADIAMQAxADkAYgA4AGYAMgA2ADgAMwBhADkANgAxAGYANgAwADMAZQA2ADQAMAAwAGEAZgBiADUAMwA5AGMAMQBhADcAZAAzAGYAYQA="
		$Key = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,42)
		$Script:mailuser = "MNSUKCATE\Y0432156"
		$Script:P_User = $serverdom.Split('.')[0] + "\" + "Y0432156"
		$Script:P_PWord = $encrypted | ConvertTo-SecureString -Key $key
		$Script:Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $P_User, $P_PWord
		$Script:mailcred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $mailuser, $P_PWord
		#$Script:Stgpath = "\\10.128.11.10\g$\NPEMS_TEMP\NPEMS\SIP\UGP\Wintel\Access_Output\ZeroTouch"
		$Script:Stgpath = "\\10.128.11.10\npems\SIP\UGP\Wintel\Access_Output\ZeroTouch"
	} catch [System.Management.Automation.RuntimeException] {
		Write-Log "We are unable to send password to user `"$pid`", since this account is not available in PROD" 
	} catch [System.IO.FileNotFoundException] {
		Write-Log "Active Directory Tools were not installed in this Machine, Please install this feature"
	} catch {
		Write-Log "Drive Exception"
		Write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		Write-Log "Exception Message: $($_.Exception.Message)"
	}
}

Function Write-Comment
{
	$cmt = $null
	$cmtsec = $commentlog -split ';'
	foreach ($sec in $cmtsec)
	{
		$cmt += " + Char(13) + " + "`'$sec`'"
	}
	$upqry = "UPDATE [$database].[dbo].[$sqltbl] SET log=$cmt WHERE ReqNo=`'$REQ`' and MemberName =`'$Membername`';"
	#write "$upqry"
	$completion = $db.ExecuteWithResults("$upqry")
	Write-Log ""
	Write-Log "Request Completed"
}

Function Revoke-Access ([string]$sw)
{
	Start-Main
	If ($sw -eq 'ADGrp')
	{
		Remove-ADGroupMember -Identity "$Groupname" -Members "$Membername" -Server $serverdom -Credential $creds -Confirm:$false
		if ($?)
		{
			Write-log "SUCCESS:$Membername`: User removed from group `"$Groupname`" in $Dcaps domain"
			$commentlog += "SUCCESS:$Membername`: User removed from group `"$Groupname`" in $Dcaps domain;"
		} else {
			Write-Log "$error[0].exception.message"
			Write-Log "`"$Membername`" : Unknown error check with admin"
			$commentlog += "`"$Membername`" : Unknown error check with admin;"
		}
	} elseif ($sw -eq 'Share')
	{
		$acl = Get-Acl "$Accpath"
		If ($acl.access | where {$_.identityreference -eq "$Currdom\$Membername"})
		{
			$newacl = "$Currdom\$Membername","Modify","ContainerInherit,ObjectInherit","None","Allow"
			$accessRule = new-object System.Security.AccessControl.FileSystemAccessRule $newacl
			$acl.RemoveAccessRule($accessRule)
			$acl | Set-Acl "$Accpath"
			if ($?)
			{
				Write-Log "Access removed to the given share path for user `"$Membername`" in $Dcaps domain"
				$commentlog += "Access removed to the given share path for user `"$Membername`" in $Dcaps domain;"
			} else {
				Write-Log "$error[0].exception.message"
				Write-Log "`"$Membername`" : Unknown error check with admin"
				$commentlog += "`"$Membername`" : Unknown error check with admin;"
			}
		} else {
			Write-Log "User `"$Membername`" does not have access to given share path in $Dcaps domain"
			$commentlog += "User `"$Membername`" does not have access to given share path in $Dcaps domain;"
		}
	}
	Write-Comment
}

Function Revoke-ZeroTouch
{
	$LogFile = "$Stgpath\$REQ`_Revoke.log"
	$result = $table = $pids = $path = $Type = $Computers = $platform = $Usernames = $Access = `
	$Category = $Workstations = $DBName = $SQLusername = $groupnames = $frdate = $todate = $null
	
	$server = "MSHSRMNSUKC0142"
	$database = "AccessManagement"
	[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
	[Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMOExtended") | out-null
	$Script:srv = new-object('Microsoft.SqlServer.Management.Smo.Server') $server
	$srv.ConnectionContext.LoginSecure = $false
	$srv.ConnectionContext.set_Login("ZeroSQL")
	$encryptedpw = "76492d1116743f0423413b16050a5345MgB8AEEAUQBhAHMAVwA5AFcAcQBQAFMAdgBtAEcAKwB6AG4AdgBuAHkAcQA5AGcAPQA9AHwAZgA1AGEAMABiADIAZQBhAGQAMwBjAGIAMwA3AGIAMQBhAGQAYwAzADkAZgA3ADUAOABhADAAOQAxADgAZAAzADEAOQBhADYAOAA3ADkANQAwADgAMwA2ADEAYQAyAGEAYQA2AGIAMAAxAGYANQBiADMAMgA3ADcAYQBiADQAZQA="
	$pwKey = (2,3,1,2,55,33,253,221,0,0,1,22,41,55,34,234,2,33,3,8,7,6,36,44)
	$SqlUPwd = $encryptedpw | ConvertTo-SecureString -Key $pwkey
	$srv.ConnectionContext.set_SecurePassword($SqlUpwd)
	$db = New-Object Microsoft.SqlServer.Management.Smo.Database
	$db = $srv.Databases.Item("$database")
	$db.ExecuteNonQuery("CHECKPOINT")
	if (($srv.Information).ComputerNamePhysicalNetBIOS -match $server)
	{
		if (($srv.Databases | ? {$_.Name -eq "$database"}).Name -Match $database)
		{	
			$SPCats = "ADGroup","Sharedpath,MNSUKCATE"
			#$SPCats = "ADGroup"
			Foreach ($SP in $SPCats)
			{
				$sqltext = "Exec revoke_access $SP;"
				$result = $db.ExecuteWithResults($sqltext);
				$table = $result.Tables[0];
				foreach ($row in $table)
				{
					$REQ = $row.Item("ReqNo")
					$DBDate = $row.Item("To")
					$Membername = $row.Item("Membername")
					$Script:domain = $row.Item("domain")
					$Script:mail = $row.Item("Email")
					if ($SP -eq "ADGroup")
					{
						$Mailip = "AD Group Mapping"
						$Sw = "ADGrp"
						$Groupname = $row.Item("groupname")
						$sqltbl = "Mapuserstogroups_test"
					} elseif ($SP -eq "Sharedpath,MNSUKCATE")
					{
						$pathtype = $row.Item("Pathtype")
						$Mailip = "$pathtype access for MNSUKCATE domain"
						$Sw = "Share"
						$Accpath = $row.Item("Path")
						$sqltbl = "SharepathAccess_test"
					}
					$ExpDate = $DBDate.Trim()
					$today = (Get-Date).date
					if ($(Get-Date $ExpDate) -lt $today)
					{
						Revoke-Access $sw
						Send-RevokeMail -Revoked
					} else {
						$today3 = ((Get-Date).Adddays(3)).date
						$today6 = ((Get-Date).Adddays(6)).date
						$today9 = ((Get-Date).Adddays(9)).date
						$today12 = ((Get-Date).Adddays(12)).date
						<#$today = "{0:dd-MMM-yy}" -f (Get-Date)
						$Today3 = "{0:dd-MMM-yy}" -f (Get-Date).Adddays(3)
						$Today6 = "{0:dd-MMM-yy}" -f (Get-Date).Adddays(6)
						$Today9 = "{0:dd-MMM-yy}" -f (Get-Date).Adddays(9)
						#>
						if (($(Get-Date $ExpDate) -eq $today) -or ($(Get-Date $ExpDate) -eq $today3) -or ($(Get-Date $ExpDate) -eq $today6) -or ($(Get-Date $ExpDate) -eq $today9) -or ($(Get-Date $ExpDate) -eq $today12))
						{							
							Send-RevokeMail -Extension
						}
						#Write-host "$Membername : $Domain : $groupname : $ExpDate"
						#break;
					}
				}
			}
		} else {
			Write-Log "DB `"$database`" not found"
		}
	} else {
		Write-Log "DB server connectivity issues"
	}
} Revoke-ZeroTouch


<#

$tot = $null
for ($i=0;$i -lt 15;$i++) {
	[string[]]$all = "{0:dd-MMM-yy}" -f ((get-date $exp).AddDays($i))
	$tot += $all
}

foreach ($ExpDate in $tot) {
	$today = (Get-Date).date
	if ($(Get-Date $ExpDate) -lt $today)
	{
		#Revoke-Access $sw
		#Send-RevokeMail -Revoked
		Write-host "Revoked: $Membername : $Domain : $groupname : $ExpDate"
	} else {
		$today3 = ((Get-Date).Adddays(3)).date
		$today6 = ((Get-Date).Adddays(6)).date
		$today9 = ((Get-Date).Adddays(9)).date
		$today12 = ((Get-Date).Adddays(12)).date
		if (($(Get-Date $ExpDate) -eq $today) -or ($(Get-Date $ExpDate) -eq $today3) -or ($(Get-Date $ExpDate) -eq $today6) -or ($(Get-Date $ExpDate) -eq $today9) -or ($(Get-Date $ExpDate) -eq $today12))
		{							
			#Send-RevokeMail -Extension
			Write-host "Extension: $Membername : $Domain : $groupname : $ExpDate"
		}
	}
}

#>