# ##################################################################################################################################
# NAME: ZeroTouch.ps1
# 
# AUTHOR:  Balaji Srinivasan (TCS)
# DATE:  2016/06/27
# EMAIL: balaji.srinivasan@marks-and-spencer.com
# 
# COMMENT:  This script will grant all types of wintel access to non-prod environment using RoD as front-end
#
# VERSION HISTORY
# 1.0 2016.04.22 Initial Version.
# 1.1 2016.05.14 Upgrade with DB logs
# 1.2 2017.01.23 Email notification for SQL user creation
# 1.3 2017.02.02 Get multiple DB access for single server
#
# ##################################################################################################################################



########################################
# Common Functions
########################################

if ((Get-ExecutionPolicy) -ne "Unrestricted"){
Set-ExecutionPolicy Unrestricted -Force
}

Function Write-Log([string]$message)
{
   Out-File -InputObject $message -FilePath $LogFile -Append
}

Function UserMgmt-SendEmail
{
	$bcc = "grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com"
	if ($Category -eq "SQL User Creation")
	{
		#[string[]]$bccaddress = "Haritha.Veerabadran@marks-and-spencer.com","Balaji.Srinivasan@marks-and-spencer.com"
		$to = $Email
		$fn = (($Email -split '@')[0]).Split('.')[0]
		$body = "Hi <b>$fn</b> <br><br>"
		$body += "Requested SQL account has been created. Please find the credentials below <br><br>"
		$body += "<u>Credentials</u>: <br>"
		$body += "<b>Username</b>: $User <br>"
		$body += "<b>Password</b>: $Password <br><br>"
		$body += "Regards,<br>Non-Prod Wintel<br><br>"
		$Body += "*** Please do not reply, This is a system generated email *** <br>"
		$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
		$sub = "$REQ`:SQL account credentials"
	} else { 
		$fn = $name.givenname -replace '\s',''
		$sn = $name.surname -replace '\s',''
		$sam = $name.SamAccountName
		$doma = $serverdom.Split('.')[0].ToString().Toupper()
		if ($sam -match "\d")
		{
			if ($fn -and $sn)
			{
				$userdata = Get-ADUser -Filter "givenname -eq '$fn' -and Surname -eq '$sn'" -Properties UserPrincipalName -Server $serverdom -Credential $creds
				if ($userdata)
				{
					[string[]] $to = $($fn + '.' + $sn + '@marks-and-spencer.com'),$Email
				} else {
					if ($email)
					{	
						$to = $Email
					}
				}
			} else {
				Write-Log "$sam`: First Name and Last Name is mandatory"
			}
		} else {
			$userdata = Get-ADUser -Filter "samaccountname -eq '$sam'" -Properties UserPrincipalName -Server $serverdom -Credential $creds
			[string[]]$to = $(($userdata.UserPrincipalName).Split('@')[0] + '@marks-and-spencer.com'),$Email
		}
		if ($Category -eq "Password Reset")
		{
			$Body = "Hi <b>$fn</b> <br><br> "
			$body += "New password has been set for your <b>$doma\$sam</b> user account <br> "
			$body += "Your new password is: <b>$Rdmpwd</b> <br><br>"
			$body += "<b>Note</b>: Try logging with your new password after 1 to 2 mins<br><br>"
			$body += "Regards,<br>Non-Prod Wintel<br><br>"
			$Body += "*** Please do not reply, This is a system generated email *** <br>"
			$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
			$sub = "$REQ`:New $doma Password"
		} elseif ($Category -eq "Account Unlock")
		{
			$body = "Hi <b>$fn</b> <br><br>"
			$body += "Your <b>$doma\$sam</b> account is unlocked <br><br>"
			$body += "<b>Note</b>: Try logging with your user account after 1 to 2 mins<br><br>"
			$body += "Regards,<br>Non-Prod Wintel<br><br>"
			$Body += "*** Please do not reply, This is a system generated email *** <br>"
			$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
			$sub = "$REQ`:$doma account is unlocked"
		}
	}
	
	try
	{
		#write-host Send-MailMessage -From "Non-Prod Access Management <Non-ProdAccessManagement@marksandspencercate.com>" -To $to -BCC "$bcc" -Subject $sub -Body $Body -BodyAsHtml -SmtpServer 10.151.209.17 -Credential $Mailcred -EA Stop
		Send-MailMessage -From "Non-Prod Access Management <Non-ProdAccessManagement@marksandspencercate.com>" -To $to -BCC "$bcc" -Subject $sub -Body $Body -BodyAsHtml -SmtpServer 10.151.209.17 -Credential $Mailcred -EA Stop
		#Send-MailMessage -From "Non-Prod User Management <Non-ProdUserManagement@marksandspencercate.com>" -To "$to@marks-and-spencer.com" -BCC "balaji.srinivasan@marks-and-spencer.com" -Subject "Your new $doma Password" -Body $Body -BodyAsHtml -SmtpServer 10.151.209.17 -Credential $Mailcred -EA Stop
	} catch [System.Net.Mail.SmtpException] {
		Write-Log "Email not sent to user $sam`: Invalid Email Address"
	} catch {
		Write-Log "Reset Mail Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
	}
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
		$Script:d = $Domain.Tostring().ToLower()
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

########################################
# DB Logs Functions
########################################

Function Database-Log ([Switch]$New, [Switch]$Extension)
{
	$Usertbl = "DBAccess"
	$Dblogdb = New-Object Microsoft.SqlServer.Management.Smo.Database
	$Dblogdb = $srv.Databases.Item("$DBlog")
	$Dblogdb.ExecuteNonQuery("CHECKPOINT")
	if (($srv.Databases | ? {$_.Name -eq "$DBlog"}).Name -Match $DBlog)
	{
		if ((($srv.Databases | ? {$_.Name -eq "$DBlog"}).tables | ? {$_.Name -eq $Usertbl}).name -match $Usertbl)
		{
			if ($new)
			{
				$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$REQ','$Computers','$DBName','$user','$access','$From','$To','$Currdom','SUCCESS','User Mapped To Database','$Email');"	
			} elseif ($Extension)
			{
				$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$REQ','$Computers','$DBName','$user','$access','$From','$To','$Currdom','SUCCESS','User Already Member','$Email');"
			}
			$result = $Dblogdb.ExecuteWithResults($sqltext);
		} else {
			Write-Log "Table `"$Usertbl`" not found"
		}
	} else {
		Write-Log "DB `"$DBlog`" not found"
	}
}

Function AddGrpDB-update ([Switch]$New, [Switch]$Extension)
{
	$Usertbl = "Mapuserstogroups"
	$Addgrpdb = New-Object Microsoft.SqlServer.Management.Smo.Database
	$Addgrpdb = $srv.Databases.Item("$DBlog")
	$Addgrpdb.ExecuteNonQuery("CHECKPOINT")
	if (($srv.Databases | ? {$_.Name -eq "$DBlog"}).Name -Match $DBlog)
	{
		if ((($srv.Databases | ? {$_.Name -eq "$DBlog"}).tables | ? {$_.Name -eq $Usertbl}).name -match $Usertbl)
		{
			if ($new)
			{
				$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$REQ','$Groupname','$user','$From','$To','$($serverdom.Split('.')[0])','SUCCESS','User Mapped To Group','$Email');"
			} elseif ($Extension)
			{
				$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$REQ','$Groupname','$user','$From','$To','$($serverdom.Split('.')[0])','SUCCESS','User Already Member','$Email');"
			}
			$result = $Addgrpdb.ExecuteWithResults($sqltext);
		} else {
			Write-Log "Table `"$Usertbl`" not found"
		}
	} else {
		Write-Log "DB `"$DBlog`" not found"
	}
}

Function UserDB-update
{
	$Usertbl = "UserGroup-Creation"
	$Userdb = New-Object Microsoft.SqlServer.Management.Smo.Database
	$userdb = $srv.Databases.Item("$DBlog")
	$userdb.ExecuteNonQuery("CHECKPOINT")
	if (($srv.Databases | ? {$_.Name -eq "$DBlog"}).Name -Match $DBlog)
	{
		if ((($srv.Databases | ? {$_.Name -eq "$DBlog"}).tables | ? {$_.Name -eq $Usertbl}).name -match $Usertbl)
		{
			$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$($serverdom.Split('.')[0])','$User','$password','$Type','SUCCESS','SQL User Created','$Email','$Req');"
			$result = $userdb.ExecuteWithResults($sqltext);
		} else {
			Write-Log "Table `"$Usertbl`" not found"
		}
	} else {
		Write-Log "DB `"$DBlog`" not found"
	}
}

Function Share-Update ([Switch]$New, [Switch]$Extension)
{
	$Usertbl = "SharepathAccess"
	$Sharedb = New-Object Microsoft.SqlServer.Management.Smo.Database
	$Sharedb = $srv.Databases.Item("$DBlog")
	$Sharedb.ExecuteNonQuery("CHECKPOINT")
	if (($srv.Databases | ? {$_.Name -eq "$DBlog"}).Name -Match $DBlog)
	{
		if ((($srv.Databases | ? {$_.Name -eq "$DBlog"}).tables | ? {$_.Name -eq $Usertbl}).name -match $Usertbl)
		{
			if ($new)
			{
				$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$REQ','$drop','$user','$From','$To','$Currdom','$Type','$PathType','SUCCESS','User Mapped To Group','$Email');"
			} elseif ($Extension)
			{
				$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$REQ','$drop','$user','$From','$To','$Currdom','$Type','$PathType','SUCCESS','User Already Member','$Email');"
			}
			$result = $Sharedb.ExecuteWithResults($sqltext);
		} else {
			Write-Log "Table `"$Usertbl`" not found"
		}
	} else {
		Write-Log "DB `"$DBlog`" not found"
	}
}

Function Write-DBlog ([Switch]$Data, [Switch]$Log)
{
	if ($data)
	{
		$LogQuery = "insert into [$DBlog].[dbo].[$sqltbl] values ('$REQ','$platform','$Category','$Domain');"
	} elseif ($Log)
	{
		$LogQuery = "insert into [$DBlog].[dbo].[$sqltbl] values ('$REQ','$platform','$Category','$Domain','$($Output.Log)');"
	}
	$ds = $db.ExecuteWithResults("$LogQuery")
}

Function Write-Comment
{
	$cmt = $null
	$cmtsec = $commentlog -split ';'
	foreach ($sec in $cmtsec)
	{
		$cmt += " + Char(13) + " + "`'$sec`'"
	}
	if ($Count -eq $totchk)
	{
		$upqry = "UPDATE [$database].[dbo].[$sqltbl] SET current_status='Completed',comments=`'$Category`: `'$cmt ,last_updated=getdate() WHERE req_id=`'$REQ`';"
		$completion = $db.ExecuteWithResults("$upqry")
		Write-Log ""
		Write-Log "Request Completed"
	} else {
		$upqry = "UPDATE [$database].[dbo].[$sqltbl] SET current_status='Completed',comments=`'Exception Occured: `'$cmt ,last_updated=getdate() WHERE req_id=`'$REQ`';"
		$completion = $db.ExecuteWithResults("$upqry")
		Write-Log ""
		Write-Log "Request Not Completed"
	}
}

########################################
# Access Functions
########################################

Function Server-Access
{
	Start-Main
    try
	{
		$end = Get-Date $todate
		$sixmon = (get-date $frdate).AddMonths(6)
		if ($end -lt $sixmon) 
		{
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		} else {
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (get-date $frdate).AddMonths(6)
		}
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Server-Access.log"
		Write-Log -Message "Server Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
        $i = 0
		$commentlog = $null
		$commentlog = @()
        $compnames = ($computers.tostring().ToUpper() -split ',').trim()
        $ccount = $compnames.count
        $users = ($usernames -split ',').trim()
        $ucount = $users.count
        $count = $ccount * $ucount
		foreach ($computer in $compnames)
		{
			if ($Computers -ne "MSHSRMNSUKC0142")
			{
				$Srvcount = $computer.ToString().length
				if ($Srvcount -eq "15")
				{
					if ($access -eq "Admin")
					{
						$Groupname = "IT SRV $Dcaps $computer Admin"
						$LocalGroup = [ADSI]"WinNT://$computer/Administrators,group"
						$OUPath = "OU=SRV Groups,OU=User Groups,OU=HO Users,OU=Head Office,DC=mnsuk$d,DC=adroot$d,DC=marksandspencer$d,DC=com"
					} elseif ($access -eq "Regular")
					{
						$Groupname = "IT $Dcaps $computer RDP"
						$LocalGroup = [ADSI]"WinNT://$computer/Remote Desktop Users,group"
						$OUPath = "OU=IT Groups,OU=User Groups,OU=HO Users,OU=Head Office,DC=mnsuk$d,DC=adroot$d,DC=marksandspencer$d,DC=com"
					}
					$DomainGroup = [ADSI]"WinNT://MNSUKCATE/$Groupname,group"
					$Grp = Get-ADGroup -Filter "Samaccountname -eq '$groupname'"
					
					if (!($grp))
					{
						New-ADGroup -SamAccountName "$Groupname" -Path "$OUPath" -Name "$Groupname" -GroupCategory "Security" -GroupScope "Global"
						if ($?)
						{
							$Grp = Get-ADGroup -Filter "Samaccountname -eq '$groupname'"
							if ($Grp)
							{
								Write-Log "SUCCESS:New AD group `"$Groupname`" Created in $Dcaps domain"
								$commentlog += "SUCCESS:New AD group `"$Groupname`" Created in $Dcaps domain;"
							} else {
								Write-Log "`"$Groupname`" Group creation failed in $Dcaps domain. Please check with Non-Prod Wintel team"
								$commentlog += "`"$Groupname`" Group creation failed in $Dcaps domain. Please check with Non-Prod Wintel team;"
								Write-Log "$error[0].exception.message"
								break;
							}
						}
					}
					
					$grpchk = ($LocalGroup.psbase.invoke('members')  | ForEach { $_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null) }) -Contains "$Groupname"
					if (!($grpchk))
					{
						$LocalGroup.Add($DomainGroup.Path)
						if ($?)
						{
							Write-Log "SUCCESS:`"$groupname`" group mapped to `"$computer`" server"
							$commentlog += "SUCCESS:`"$groupname`" group mapped to `"$computer`" server;"
						} else {
							Write-Log "Error while adding domain group `"$groupname`" to local group($access)"
							$commentlog += "Error while adding domain group `"$groupname`" to local group($access);"
							Write-Log "$error[0].exception.message"
							break;
						}
					} else {
						Write-Log "SUCCESS:`"$groupname`" group already mapped to `"$computer`" server"
						$commentlog += "SUCCESS:`"$groupname`" group already mapped to `"$computer`" server;"
					}
					
					Foreach ($user in $Users)
					{
						if ($user -match "\.")
						{
							$mem = Get-ADUser -Filter "Samaccountname -eq '$user'" -Properties MemberOf
							if ($mem)
							{
								$Testmem = $mem.MemberOf | foreach {$_ -eq $Grp.DistinguishedName} | where {$_ -eq $true}
								if (!($Testmem))
								{
									if ($Type -eq "New")
									{	
										Add-ADGroupMember -Identity "$groupname" -Members "$user"
										if ($?)
										{
											AddGrpDB-update -New
											Write-log "SUCCESS:`"$user`" : User mapped to server AD group `"$groupname`" in $Dcaps domain"
											$commentlog += "SUCCESS:`"$user`" : User mapped to server AD group `"$groupname`" in $Dcaps domain;"
											$totchk = ++$i
										} else {
											Write-Log "Error while adding user `"$user`" to domain group `"$groupname`""
											$commentlog += "Error while adding user `"$user`" to domain group `"$groupname`";"
											Write-Log "$error[0].exception.message"
										}
									} elseif ($Type -eq "Extension")
									{
										Write-Log "`"$user`" : User does not have access to `"$groupname`" AD group in $Dcaps domain. Please select request type as NEW"
										$commentlog += "`"$user`" : User does not have access to `"$groupname`" AD group in $Dcaps domain. Please select request type as NEW;"
									}
								} else {
									AddGrpDB-update -Extension
									Write-Log "SUCCESS:`"$user`" : User already member to `"$groupname`" AD group in $Dcaps domain"
									$commentlog += "SUCCESS:`"$user`" : User already member to `"$groupname`" AD group in $Dcaps domain;"
									$totchk = ++$i
								}
							} else {
								Write-Log "User `"$user`" not found in $Dcaps domain"
								$commentlog += "User `"$user`" not found in $Dcaps domain;"
							}
						} else {
							Write-Log "`"$user`" : This is not a Name.Name Account"
							$commentlog += "`"$user`" : This is not a Name.Name Account;"
						}
					}
				} else {
					Write-Log "`"$Computer`" : Given computer is not a server"
					$commentlog += "`"$Computer`" : Given computer is not a server;"
				}
			} else {
				Write-log "`"$computers`" : Access will not be provided to this server, since it is non-prod server"
				$commentlog += "`"$computers`" : Access will not be provided to this server, since it is non-prod server;"
			}
		}
	} catch {
		Write-Log "Server Access Exception"
        write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
    }
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Workstation-Access
{
	try
	{
		Start-Main
		$end = Get-Date $todate
		$sixmon = (get-date $frdate).AddMonths(6)
		if ($end -lt $sixmon) 
		{
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		} else {
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (get-date $frdate).AddMonths(6)
		}
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Application.log"
		Write-Log -Message "Workstation Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
		$i = 0
		$commentlog = $null
		$commentlog = @()
        $wks = ($Workstations.tostring().ToUpper() -split ',').trim()
        $ccount = $wks.count
        $users = ($Usernames -split ',').trim()
        $ucount = $users.count
        $count = $ccount * $ucount
		foreach ($wk in $wks)
		{
			Write-Log "Workstation: `"$wk`""
			$Wkcount = $wk.ToString().length
			if ($Wkcount -ne "15")
			{
				$Groupname = "IT $Dcaps $wk"
				$Grp = Get-ADGroup -Filter "Samaccountname -eq '$groupname'" -Server $serverdom -Credential $creds
				if ($grp)
				{
					Foreach ($user in $Users)
					{
						if ($user -match "P\d")
						{
							$mem = Get-ADUser -Filter "Samaccountname -eq '$user'" -Server $serverdom -Credential $creds -Properties memberof
							if ($mem)
							{
								$Testmem = $mem.MemberOf | foreach {$_ -eq $Grp.DistinguishedName} | where {$_ -eq $true}
								if (!($Testmem))
								{
									if ($Type -eq "New")
									{	
										Add-ADGroupMember -Identity "$groupname" -Members "$user" -Server $serverdom -Credential $creds
										if ($?)
										{
											AddGrpDB-update -New
											Write-log "SUCCESS:`"$user`" : User mapped to group `"$groupname`" in $Dcaps domain"
											$commentlog += "SUCCESS:`"$user`" : User mapped to group `"$groupname`" in $Dcaps domain;"
											$totchk = ++$i
										} else {
											Write-Log "$error[0].exception.message"
											Write-Log "`"$user`" : Unknown error check with admin"
											$commentlog += "`"$user`" : Unknown error check with admin;"
										}
									} elseif ($Type -eq "Extension")
									{
										Write-Log "`"$user`" : User does not have access to group `"$groupname`" in $Dcaps domain. Please select request type as NEW"
										$commentlog += "`"$user`" : User does not have access to group `"$groupname`" in $Dcaps domain. Please select request type as NEW;"
									}
								} else {
									AddGrpDB-update -Extension
									Write-Log "SUCCESS:`"$user`" : User already member to group `"$groupname`" in $Dcaps domain"
									$commentlog += "SUCCESS:`"$user`" : User already member to group `"$groupname`" in $Dcaps domain;"
									$totchk = ++$i
								}
							} else {
								Write-Log "`"$user`" : User not found in $Dcaps domain"
								$commentlog += "`"$user`" : User not found in $Dcaps domain;"
							}
						} else {
							Write-Log "`"$user`" : This is not a PID user"
							$commentlog += "`"$user`" : This is not a PID user;"
						}
					}
				} else {
					Write-Log "Group `"$groupname`" not found in $Dcaps domain. Please check with wintel team for validation"
					$commentlog += "Group `"$groupname`" not found in $Dcaps domain. Please check with wintel team for validation;"
				}
			} else {
				Write-Log "Given computername `"$wk`" is not a workstation"
				$commentlog += "Given computername `"$wk`" is not a workstation;"
			}
		}
	} catch {
		Write-Log "Workstation Access Exception"
        write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
    }
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Set-Application
{
	Start-Main
	try
	{
		$end = Get-Date $todate
		$sixmon = (get-date $frdate).AddMonths(6)
		if ($end -lt $sixmon) 
		{
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		} else {
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (get-date $frdate).AddMonths(6)
		}
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Application.log"
		Write-Log -Message "Application Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
        $Users = ($Pids -split ',').trim()
		$Groups = ($Groupnames -split ',').trim()
        $ucount = $Users.count
		$gcount = $Groups.count
		$count = $gcount * $ucount
		foreach ($Groupname in $groups)
		{
			Write-Log "GroupName: `"$Groupname`""
			$commentlog += "GroupName: `"$Groupname`""
			if (($Groupname -eq "IT APP XP Advanced Users") -or ($Groupname -like "IT SRV*") -or ($Groupname -like "IT ADM*"))
			{
				Write-Log "`"$Groupname`" : This group required SPAF approval. Please check with Non-Prod Wintel team"
				$commentlog += "`"$Groupname`" : This group required SPAF approval. Please check with Non-Prod Wintel team;"
			} else {
				$Grp = Get-ADGroup -Filter "Samaccountname -eq '$Groupname'" -Server $serverdom -Credential $creds
				if ($grp)
				{
					foreach ($User in $Users)
					{
						if ($user -match "\d")
						{
							$mem = Get-ADUser -Filter "Samaccountname -eq '$user'" -Server $serverdom -Credential $creds -Properties memberof
							if ($mem)
							{
								$Testmem = $mem.MemberOf | foreach {$_ -eq $Grp.DistinguishedName} | where {$_ -eq $true}
								if (!($Testmem))
								{
									if ($Type -eq "New")
									{	
										Add-ADGroupMember -Identity "$Groupname" -Members "$user" -Server $serverdom -Credential $creds
										if ($?)
										{
											AddGrpDB-update -New
											Write-log "SUCCESS:$user`: User mapped to group `"$Groupname`" in $Dcaps domain"
											$commentlog += "SUCCESS:$user`: User mapped to group `"$Groupname`" in $Dcaps domain;"
											$totchk = ++$i
										} else {
											Write-Log "$error[0].exception.message"
											Write-Log "`"$user`" : Unknown error check with admin"
											$commentlog += "`"$user`" : Unknown error check with admin;"
										}
									} elseif ($Type -eq "Extension")
									{
										Write-Log "`"$user`" : User does not have access to group `"$Groupname`" in $Dcaps domain. Please select request type as NEW"
										$commentlog += "`"$user`" : User does not have access to group `"$Groupname`" in $Dcaps domain. Please select request type as NEW;"
									}
								} else {
									AddGrpDB-update -Extension
									Write-Log "SUCCESS:$user`: User already member to group `"$Groupname`" in $Dcaps domain"
									$commentlog += "SUCCESS:$user`: User already member to group `"$Groupname`" in $Dcaps domain;"
									$totchk = ++$i
								}
							} else {
								Write-Log "`"$user`" : User not found in $Dcaps domain"
								$commentlog += "`"$user`" : User not found in $Dcaps domain;"
							}
						} else {
							Write-Log "`"$user`" : This is not a PID user"
							$commentlog += "`"$user`" : This is not a PID user;"
						}
					}
				} else {
					Write-Log "`"$Groupname`" : Group not found in $Dcaps domain"
					$commentlog += "`"$Groupname`" : Group not found in $Dcaps domain;"
				}
			}
		}
	} catch {
		Write-Log "Application Access Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Set-ControlM
{
	Start-Main
	try
	{
		$end = Get-Date $todate
		$sixmon = (get-date $frdate).AddMonths(6)
		if ($end -lt $sixmon) 
		{
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		} else {
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (get-date $frdate).AddMonths(6)
		}
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Control-M.log"
		Write-Log -Message "Control-M Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
        $Users = ($Pids -split ',').trim()
        $count = $Users.count
		if ($Domain -eq "DEV")
		{
			$Groupname = "IT APP CTM NonProd Admin"
		} elseif ($Domain -eq "CATE")
		{
			$Groupname = "IT APP CTM Application Support"
		}
		$Grp = Get-ADGroup -Filter "Samaccountname -eq '$groupname'" -Server $serverdom -Credential $creds
		foreach ($User in $Users)
		{
			if ($user -match "P\d")
			{
				$mem = Get-ADUser -Filter "Samaccountname -eq '$user'" -Server $serverdom -Credential $creds -Properties memberof
				if ($mem)
				{
					$Testmem = $mem.MemberOf | foreach {$_ -eq $Grp.DistinguishedName} | where {$_ -eq $true}
					if (!($Testmem))
					{
						if ($Type -eq "New")
						{
							Add-ADGroupMember -Identity "$Groupname" -Members "$user" -Server $serverdom -Credential $creds
							if ($?)
							{
								AddGrpDB-update -New
								Write-log "SUCCESS:$user`: Access provided to $Dcaps Control-M"
								$commentlog += "SUCCESS:$user`: Access provided to $Dcaps Control-M;"
								$totchk = ++$i
							} else {
								Write-Log "$error[0].exception.message"
								Write-Log "`"$user`" : Unknown error check with admin"
								$commentlog += "`"$user`" : Unknown error check with admin;"
							}
						} elseif ($Type -eq "Extension")
						{
							Write-Log "`"$user`" : User does not have access to $Dcaps Control-M. Please select request type as NEW"
							$commentlog += "`"$user`" : User does not have access to $Dcaps Control-M. Please select request type as NEW;"
						}
					} else {
						AddGrpDB-update -Extension
						Write-Log "SUCCESS:$user`: User already has access to $Dcaps Control-M. Account Extended"
						$commentlog += "SUCCESS:$user`: User already has access to $Dcaps Control-M. Account Extended;"
						$totchk = ++$i
					}
				} else {
					Write-Log "`"$user`" : User not found in $Dcaps domain"
					$commentlog += "`"$user`" : User not found in $Dcaps domain;"
				}
			} else {
				Write-Log "`"$user`" : Access will be given only to PID user"
				$commentlog += "`"$user`" : Access will be given only to PID user;"
			}
		}
	} catch {
		Write-Log "Control-M Access Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Set-StagingPath
{
	Start-Main
	try
	{
		$end = Get-Date $todate
		$sixmon = (get-date $frdate).AddMonths(6)
		if ($end -lt $sixmon) 
		{
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		} else {
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (get-date $frdate).AddMonths(6)
		}
		$LogFile = "$Stgpath" + "\$REQ.log"
		$PathType = "Staging Path"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Staging-Access.log"
		Write-Log -Message "Staging Access: $Today`: $REQ - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
		$Paths = ($Pathlist -split ',').trim()
		$Users = ($pids -split ',').trim()
		$ucount = $Users.count
		$pcount = $Paths.count
		$count = $ucount * $pcount
		foreach ($Path in $Paths)
		{
			Write-Log -Message "Given Staging Path: $Path"
			$commentlog += "Given Staging Path: $Path"
			if (($Path[0..26] -join "") -eq "\\10.128.11.10\Nongustaging")
			{
				$drop = $null
				if ($Path.tostring().tolower().contains('drop'))
				{
					#$Testing = Split-Path $path
					#$hello.Substring($hello.Length-4)
					$Pathsplit = ($Path -split 'drop')[0]
					if (Test-path $Pathsplit\drop)
					{
						$drop = $Pathsplit + "drop"
					} else {
						Write-Log "Drop path not found. Please provide drop path to grant access"
						$commentlog += "Drop path not found. Please provide drop path to grant access;"
					}
				} elseif ($Path.tostring().tolower().contains('release')) {
					$Pathsplit = ($Path -split 'release')[0]
					if (Test-path $Pathsplit\drop)
					{
						$drop = $Pathsplit + "drop"
					} else {
						Write-Log "Drop path not found. Please provide drop path to grant access"
						$commentlog += "Drop path not found. Please provide drop path to grant access;"
					}
				} else {
					if (Test-path $Path\drop)
					{
						if ($path[-1] -eq "\")
						{
							$drop = $path + "drop"
						} else {
							$drop = $path + "\drop"
						}
					} else {
						if (Test-path $path)
						{
							$newpath = New-Item -type Directory -Path $path\drop -ErrorAction SilentlyContinue | Select -Expand Fullname
							$drop = $newpath
						} else {
							Write-Log "Given staging path is not valid"
							$commentlog += "Given staging path is not valid;"
						}
					}
				}
				if ($drop)
				{
					foreach ($user in $users)
					{
						$chkstguser = Get-ADUser -Filter "samaccountname -eq '$User'" # -Credential $Creds -Server $serverdom
						if ($chkstguser)
						{
							$acl = Get-Acl "$drop"
							If (!($acl.access | where {$_.identityreference -eq "$Currdom\$User"}))
							{
								if ($Type -eq "New")
								{
									$newacl = "$Currdom\$User","Modify","ContainerInherit,ObjectInherit","None","Allow"
									$accessRule = new-object System.Security.AccessControl.FileSystemAccessRule $newacl
									$acl.SetAccessRule($accessRule)
									$acl | Set-Acl "$drop"
									if ($?)
									{
										Share-Update -New
										Write-Log "Access provided to given drop path for user `"$User`""
										$commentlog += "Access provided to given drop path for user `"$User`";"
										$totchk = ++$i
									} else {
										Write-Log "$error[0].exception.message"
										Write-Log "`"$user`" : Unknown error check with admin"
										$commentlog += "`"$user`" : Unknown error check with admin;"
									}
								} elseif ($Type -eq "Extension")
								{
									Write-Log "User `"$User`" does not have access to given drop path"
									$commentlog += "User `"$User`" does not have access to given drop path;"
								}
							} else {
								Share-Update -Extension
								Write-Log "Access extended to given drop path for user `"$User`""
								$commentlog += "Access extended to given drop path for user `"$User`";"
								$totchk = ++$i
							}
						} else {
							Write-Log "`"$user`" : User not found in CATE domain"
							$commentlog += "`"$user`" : User not found in CATE domain;"
						}
					}
				}
			} else {
				Write-Log "Invalid staging share path"
				$commentlog += "Invalid staging share path;"
			}
		}
	} catch {
		Write-Log "Staging Path Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Set-ApplicationPath
{
	Start-Main
	try
	{
		$end = Get-Date $todate
		$sixmon = (get-date $frdate).AddMonths(6)
		if ($end -lt $sixmon) 
		{
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		} else {
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (get-date $frdate).AddMonths(6)
		}
		$PathType = "Server Path"
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\ApplicationPath-Access.log"
		Write-Log -Message "Application Path Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
		$Paths = ($pathlist -split ',').trim()
		$Users = ($pids -split ',').trim()
		$ucount = $Users.count
		$pcount = $Paths.count
		$count = $ucount * $pcount
		foreach ($Path in $paths)
		{
			Write-Log -Message "Given Share Path: $Path"
			$commentlog += "Given Share Path: $Path;"
			if ($path[12] -eq 'C')
			{	
				if (Test-path $path)
				{
					if ([string]::Join('\', $path.Split('\')[2..$($path.Split('\').Length-3)]) -ne 'MSHSRMNSUKC0142')
					{
						$drop = $Path
						foreach ($user in $users)
						{
							$chkobj = Get-ADObject -Filter "samaccountname -eq '$user'" | select -expand objectclass
							if ($chkobj)
							{
								if ($chkobj -eq "User")
								{
									$chksrvuser = Get-ADUser -Filter "samaccountname -eq '$User'"
								} elseif ($chkobj -eq "Group")
								{
									$chksrvuser = Get-ADGroup -Filter "samaccountname -eq '$User'"
								}
								if ($chksrvuser)
								{
									$acl = Get-Acl "$drop"
									If (!($acl.access | where {$_.identityreference -eq "$Currdom\$User"}))
									{
										if ($Type -eq "New")
										{
											$newacl = "$Currdom\$User","Modify","ContainerInherit,ObjectInherit","None","Allow"
											$accessRule = new-object System.Security.AccessControl.FileSystemAccessRule $newacl
											$acl.SetAccessRule($accessRule)
											$acl | Set-Acl "$drop"
											if ($?)
											{
												Share-Update -New
												Write-Log "Access provided to given share path for user `"$User`""
												$commentlog += "Access provided to given share path for user `"$User`";"
												$totchk = ++$i
											} else {
												Write-Log "$error[0].exception.message"
												Write-Log "`"$user`" : Unknown error check with admin"
												$commentlog += "`"$user`" : Unknown error check with admin;"
											}
										} elseif ($Type -eq "Extension")
										{
											Write-Log "User `"$User`" does not have access to given share path"
											$commentlog += "User `"$User`" does not have access to given share path;"
										}
									} else {
										Share-Update -Extension
										Write-Log "Access extended to given share path for user `"$User`""
										$commentlog += "Access extended to given share path for user `"$User`";"
										$totchk = ++$i
									}
								} else {
									Write-Log "`"$user`": User not found in $Dcaps domain"
									$commentlog += "`"$user`": User not found in $Dcaps domain;"
								}
							} else {
								Write-Log "`"$user`": Provided AD object is invalid"
								$commentlog += "`"$user`": Provided AD object is invalid;"
							}
						}
					} else {
						Write-log "`"$computers`" : Access will not be provided to this server, since it is non-prod dedicated server"
						$commentlog += "`"$computers`" : Access will not be provided to this server, since it is non-prod dedicated server;"
					}
				} else {
					Write-Log "Given share path not found"
					$commentlog += "Given share path not found;"
				}
			} else {
				Write-Log "Given share path is not a CATE server"
				$commentlog += "Given share path is not a CATE server;"
			}
		}
	} catch {
		Write-Log "Server Path Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Set-AccountUnlock
{
	Start-Main
	try 
	{
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Set-AccountUnlock.log"
		Write-Log -Message "Account Unlock: $Today`: $REQ ($Dcaps)"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
		$Users = ($Pids -split ',').trim()
		$count = $Users.count
		foreach ($user in $Users)
		{
			$name = Get-ADUser -Filter 'samaccountname -eq $user' -Credential $Creds -Server $serverdom -Properties LockedOut, AccountExpirationDate
			if ($name)
			{
				if ($name.LockedOut -eq $true)
				{
					Unlock-ADAccount -Server $serverdom -Identity "$user" -Credential $Creds
					if($?)
					{
						UserMgmt-SendEmail
						Write-Log -Message "SUCCESS:$User`: Account Unlocked"
						$commentlog += "SUCCESS:$User`: Account Unlocked;"
						$totchk = ++$i
					} else {
						Write-Log "$user`: $error[0].exception.message"
						$commentlog += "`"$user`" : $error[0].exception.message;"
						Write-Log "`"$user`" : Unknown error check with admin"
						$commentlog += "`"$user`" : Unknown error check with admin;"
					}
				} else {
					Write-Log -Message "$User`: Account not locked"
					$commentlog += "$User`: Account not locked;"
					$totchk = ++$i
				}
			} elseif ($name -eq $null) {
				Write-Log -Message "$User`: User not found"
				$commentlog += "$User`: User not found;"
			}
		}
	} catch {
		Write-Log "Unlock Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Set-Password
{
	Start-Main
	try
	{
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Set-Password.log"
		Write-Log -Message "Password Reset: $Today`: $REQ ($Dcaps)"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
        $Users = ($Pids -split ',').trim()
        $count = $Users.count		
		foreach ($User in $Users)
		{			
			$Random = Get-Random 100 -Minimum 10
			$arr = "london","kingdom","default","welcome","wisdom","happy","hello","secret","master","private","forever","spencer","destiny","summer"
			$Ranstr = Get-Random $arr
			$Rdmpwd = $Ranstr + $Random
			$pwd = ConvertTo-SecureString $Rdmpwd -AsPlainText -Force;
			$name = Get-ADUser -Filter "samaccountname -eq '$User'" -Credential $Creds -Server $serverdom -Properties LockedOut, AccountExpirationDate, Enabled
			if ($name -eq $null)
			{
				Write-Log "$User : User Not Found"
				$commentlog += "$User : User Not Found;"
			} else {
				if ($name.enabled)
				{
					#if (($(get-date $name.AccountExpirationDate)) -gt ($(Get-date)))
					#{
						if ($($name.AccountExpirationDate) -lt $(Get-Date))
						{
							Write-Log "WARNING: $user account is expired, raise an extension request in non-prod RoD to get it fixed"
							$commentlog += "WARNING:$user account is expired, raise an extension request in non-prod RoD to get it fixed;"
						}
						if ($name.LockedOut)
						{
							Unlock-ADAccount -Server $serverdom -Identity "$User" -Credential $Creds
						}
						Set-ADAccountPassword -Server $serverdom -Identity "$User" -Reset -NewPassword $pwd -Credential $Creds
						if($?)
						{
							UserMgmt-SendEmail
							Write-Log  "SUCCESS:Login password for user $User`: $Rdmpwd"
							$commentlog += "SUCCESS:Password reset completed for user $User`. Credentials will be sent to user in separate email;"
							$totchk = ++$i
						} else {
							Write-Log "$error[0].exception.message"
							Write-Log "`"$user`" : Unknown error check with admin"
							$commentlog += "`"$user`" : Unknown error check with admin;"					
						}
					<#} else {
						Write-Log "$user`: Account Expired, Please raise a extension request"
						$commentlog += "$user`: Account Expired, Please raise a extension request;"
					}#>
				} else {
					Write-Log "$user`: Account Disabled"
					$commentlog += "$user`: Account Disabled;"
				}
			}
		}
	} catch {
		Write-Log "Reset Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Set-DBAccess
{
	Start-Main
	try
	{
		$end = Get-Date $todate
		$sixmon = (get-date $frdate).AddMonths(6)
		if ($end -lt $sixmon) 
		{
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		} else {
			$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
			$To = "{0:dd-MMM-yy}" -f (get-date $frdate).AddMonths(6)
		}
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\DBAccess.log"
		Write-Log -Message "Database Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
        $Users = ($SQLusername -split ',').trim()
		$DBs = ($DBNames -split ',').trim()
        $ucount = $Users.count
		$dcount= $DBs.count
		$count = $ucount * $dcount
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
		if ($Computers -ne "MSHSRMNSUKC0142")
		{
			$Svr = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $Computers
			if ($Access -eq "Admin")
			{
				$Role = "db_owner"
			} elseif ($Access -eq "Regular")
			{
				$Role = "db_datawriter"
			}
			foreach ($User in $Users)
			{
				$Find = [ADSISearcher]"(sAMAccountName=$User)"
				$Obj = $Find.FindOne()
				if ($Obj -eq $null)
				{
					Write-Log "Unable to find `"$User`" on `"$Currdom`" domain"
					$commentlog += "Unable to find `"$User`" on `"$Currdom`" domain;"
				} else {
					if(!($svr.Logins.Contains("$Currdom\$User")))
					{
						$login = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Login -ArgumentList $Svr, "$Currdom\$User"
						$login.LoginType = "WindowsUser"
						$login.Create()
						if ($?)
						{
							Write-log "User `"$user`" added to server `"$computers`""
							$commentlog += "User `"$user`" added to server `"$computers`" ;"
						} else {
							Write-Log "`"$user`" : Unknown error check with admin"
							$commentlog += "`"$user`" : Unknown error check with admin;"
							Write-Log "$user`:$error[0].exception.message"
							$commentlog += "$user`:$error[0].exception.message;"
						}
					}
					foreach ($DBName in $DBs)
					{
						$dbase = $svr.Databases[$DBName]
						if ($dbase -eq $null)
						{
							Write-Log "Unable to find `"$DBName`" Database on server `"$Computers`""
							$commentlog += "Unable to find `"$DBName`" Database on server `"$Computers`" ;"
						} else {
							if (!($dbase.Users.Contains("$Currdom\$User")))
							{
								if ($type -eq "New")
								{
									$usr = New-Object ('Microsoft.SqlServer.Management.Smo.User') ($dbase, "$Currdom\$User")
									$usr.Login = "$Currdom\$User"
									$usr.Create()
									if ($?)
									{
										$Rol = $dbase.Roles[$Role]
										$Rol.AddMember("$Currdom\$User")
										if ($?)
										{
											Database-Log -New
											Write-Log "SUCCESS:Wintel Logon`:`"$User`" provided access to `"$DBName`" database as `"$Role`" on Server `"$Computers`""
											$commentlog += "SUCCESS:Wintel Logon`:`"$User`" provided access to `"$DBName`" database as `"$Role`" on Server `"$Computers`" ;"
											$totchk = ++$i
										} else {
											Write-Log "`"$user`" : Unknown error check with admin"
											$commentlog += "`"$user`" : Unknown error check with admin;"
											Write-Log "$user`:$error[0].exception.message"
											$commentlog += "$user`:$error[0].exception.message;"
										}
									} else {
										Write-Log "User `"$user`" not created on server `"$computers`""
										$commentlog += "User `"$user`" not created on server `"$computers`" ;"
										Write-Log "$user`:$error[0].exception.message"
										$commentlog += "$user`:$error[0].exception.message;"
									}
								} elseif ($type -eq "Extension")
								{
									Write-Log "User `"$user`" does not have access to database `"$DBName`". Raise a request with type NEW"
									$commentlog += "User `"$user`" does not have access to database `"$DBName`". Raise a request with type NEW;"
								}
							} else {
								$Exsist = ($dbase.Users["$Currdom\$User"]).enumroles() -match $Role
								if ($Exsist)
								{
									Database-Log -Extension
									Write-Log "SUCCESS:Wintel Logon`:`"$User`" Extended access to `"$DBName`" database as `"$Role`" on Server `"$Computers`""
									$commentlog += "SUCCESS:Wintel Logon`:`"$User`" Extended access to `"$DBName`" database as `"$Role`" on Server `"$Computers`" ;"
									$totchk = ++$i
								}
								#Import-Csv "$Filepath" | ConvertTo-Csv -NoTypeInformation | select -Skip 2 | Out-File \\10.128.11.10\npems\SIP\UGP\Wintel\Access_Output\DBAccessAuditLog.csv -Append -Encoding ascii
								#Write-Output $Output | Export-Csv \\10.128.11.10\npems\SIP\UGP\Wintel\Access_Output\DBAccessAuditLog.csv -NoTypeInformation -Append
							}
						}
					}
				}
			}
		} else {
			Write-log "`"$computers`" : Access will not be provided to this server, since it is non-prod dedicated server"
			$commentlog += "`"$computers`" : Access will not be provided to this server, since it is non-prod dedicated server;"
		}
	} catch {
		Write-Log "Database Access Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function New-SQLUser
{
	Start-Main
	try
	{
		$LogFile = "$Stgpath\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\SQLUser-Creation.log"
		Write-Log -Message "SQL User Creation: $Today`: $REQ ($Dcaps)"
		Write-Log -Message "=========================================================================================="
		$commentlog = $null
		$commentlog = @()
		$i = 0
		$Users = ($SQLusername -split ',').Trim()
		$count = $Users.count
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
		if ($Computers -ne "MSHSRMNSUKC0142")
		{
			$Svr = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $Computers
			if ($Access -eq "Admin")
			{
				$Role = "db_owner"
			} elseif ($Access -eq "Regular")
			{
				$Role = "db_datareader"
			}
			foreach ($user in $users)
			{
				$Suffix = ([char[]](Get-Random -Input $(48..57 + 65..90 + 97..122) -Count 2)) -join ""
				$pass = $User[0..4] -join ""
				$password = $pass + $suffix
				if(!($svr.Logins.Contains($User)))
				{
					$login = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Login -ArgumentList $Svr, $User
					$Login.LoginType = [Microsoft.SqlServer.Management.Smo.LoginType]::SqlLogin
					$Login.PasswordPolicyEnforced = $false
					#$Login.PasswordExpirationEnabled = $false
					$Login.Create($Password)
					if ($?)
					{
						Write-log "SQL User `"$user`" created and added to server `"$computers`""
						$commentlog += "SQL User `"$user`" created and added to server `"$computers`" ;"
					} else {
						Write-Log "`"$user`" : Unknown error check with non-prod wintel team"
						$commentlog += "`"$user`" : Unknown error check with non-prod wintel team;"
						Write-Log "$user`:$error[0].exception.message"
						$commentlog += "`"$user`" : $error[0].exception.message;"
					}
				}
				$dbase = $svr.Databases[$DBName]
				if ($dbase -eq $null)
				{
					Write-Log "Unable to find `"$DBName`" Database on server `"$computers`""
					$commentlog += "Unable to find `"$DBName`" Database on server `"$computers`" ;"
				} else {
					if (!($dbase.Users.Contains($User)))
					{
						$usr = New-Object ('Microsoft.SqlServer.Management.Smo.User') ($dbase, $User)
						$usr.Login = $User
						$usr.Create()
						if ($?)
						{
							Write-log "SQL User `"$user`" added to database `"$dbase`" on server `"$computers`""
							$commentlog += "SQL User `"$user`" added to database `"$dbase`" on server `"$computers`" ;"
						} else {
							Write-Log "`"$user`" : Unknown error check with non-prod wintel team"
							$commentlog += "`"$user`" : Unknown error check with non-prod wintel team;"
							Write-Log "$user`:$error[0].exception.message"
							$commentlog += "`"$user`" : $error[0].exception.message;"
						}
					}
					$Rol = $dbase.Roles[$Role]
					if ($Rol -eq $null)
					{
						Write-Log "DB role `"$role`" not found in `"$DBName`" database"
						$commentlog += "DB role `"$role`" not found in `"$DBName`" database;"
					} else {
						$Rol.AddMember($User)
						if ($?)
						{
							UserDB-update
							UserMgmt-SendEmail
							Write-Log "SQL Logon`:`"$User`" provided access to `"$DBName`" database as `"$Role`" on Server `"$computers`""
							Write-log "Login Credentials:	"				
							Write-log "Username: $User "
							Write-Log "Password: $Password "						
							$commentlog += "SQL Logon`:`"$User`" provided access to `"$DBName`" database as `"$Role`" on Server `"$computers`" ;"
							$commentlog += "Login credentials will be shared to requestor in separate email ;"
							$totchk = ++$i
							#Write-Output $Output | Export-Csv \\10.128.11.10\npems\SIP\UGP\Wintel\Access_Output\DBAccessAuditLog.csv -NoTypeInformation -Append
						} else {
							Write-Log "`"$user`" : Unknown error check with non-prod wintel team"
							$commentlog += "`"$user`" : Unknown error check with non-prod wintel team;"
							Write-Log "$user`:$error[0].exception.message"
							$commentlog += "`"$user`" : $error[0].exception.message;"
						}
					}
				}
			}
		} else {
			Write-log "`"$computers`" : Access will not be provided to this server, since it is non-prod dedicated server"
			$commentlog += "`"$computers`" : Access will not be provided to this server, since it is non-prod dedicated server;"
		}
	} catch {
		Write-Log "SQL User Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function Test-Function
{
	$Users = $Pids -split ','
	$count = $Users.count
	foreach ($User in $Users)
	{
		Write-host "$User"
	}
}

Function AIX-Access
{	
	#$Commentlog = C:\Wintel-SIP\Zero-Touch\plink.exe -ssh -l zerouser -pw zerouser zerouser@10.128.178.8 "sh wrapper_final.sh -s $Computers -u $Pids -g $AixGroup -e $todate -r $REQ -f $PathList -a $Access -t $Category"
	#Write-host C:\Wintel-SIP\Zero-Touch\plink.exe -ssh -l zerouser -pw zerouser zerouser@10.128.178.8 "sh wrapper_final.sh -s $Computers -u $Pids -g $AixGroup -e $todate -r $REQ -f $PathList -a $Access -t $Category -m $aixmail"
 	#$Commentlog = C:\Wintel-SIP\Zero-Touch\plink.exe -ssh -l zerouser -pw zerouser zerouser@10.128.178.8 "sh wrapper_final.sh -s $Computers -u $Pids -g $AixGroup -e $todate -r $REQ -f $PathList -a $Access -t $Category -m $aixmail"
	$Commentlog = C:\Wintel-SIP\Zero-Touch\plink.exe -ssh -l zerouser -pw zerouser zerouser@10.128.178.8 "sh wrapper_final.sh -s $Computers -u $Pids -g $AixGroup -e $todate -r $REQ -f $PathList -a $Access -t $Category"
	if ($Commentlog)
	{
		$cmt = $null
		$cmtsec = $commentlog -split ';'
		foreach ($sec in $cmtsec)
		{
			$cmt += " + Char(13) + " + "`'$sec`'"
		}
		$upqry = "UPDATE [ZeroTouchDB].[dbo].[tbl_zero_touch_access] SET current_status='Completed',comments=`'$Category`: `'$cmt ,last_updated=getdate() WHERE req_id=`'$REQ`';"
		$completion = $db.ExecuteWithResults("$upqry")
	} else {
		$upqry = "UPDATE [ZeroTouchDB].[dbo].[tbl_zero_touch_access] SET current_status='Completed',comments=`'Exception Occured: Please check with Non-Prod AIX team`' ,last_updated=getdate() WHERE req_id=`'$REQ`';"
		$completion = $db.ExecuteWithResults("$upqry")
	}
	
}

########################################
# Root Function
########################################

Function Complete-ZeroTouch
{
	$result = $table = $pids = $path = $Type = $Computers = $platform = $Usernames = $Access = `
	$Category = $Workstations = $DBName = $SQLusername = $groupnames = $frdate = $todate = $null
	<#[void][reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")
	[void][reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo")
	$server = "MSHSRMNSUKC0142"
	$database = "ZeroTouchDB"
	$srv = new-Object Microsoft.SqlServer.Management.Smo.Server("$server")
	$db = New-Object Microsoft.SqlServer.Management.Smo.Database
	$db = $srv.Databases.Item($database)
	$db.ExecuteNonQuery("CHECKPOINT")
	$sqltext = "Select * from [ZeroTouchDB].[dbo].[tbl_zero_touch_access] Where current_status = 'ImplementationInProgress';"
	$result = $db.ExecuteWithResults($sqltext);
	$table = $result.Tables[0];#>
	
	$server = "MSHSRMNSUKC0142"
	$database = "ZeroTouchDB"
	$SQLtbl = "tbl_zero_touch_access"
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
			if ((($srv.Databases | ? {$_.Name -eq "$database"}).tables | ? {$_.Name -eq $sqltbl}).name -match $sqltbl)
			{
				$sqltext = "Select * from [$database].[dbo].[$SQLtbl] Where current_status = 'ImplementationInProgress';"
				$result = $db.ExecuteWithResults($sqltext);
				$table = $result.Tables[0];
			} else {
				Write-Log "Table `"$SQLtbl`" not found"
			}
		} else {
			Write-Log "DB `"$database`" not found"
		}
	} else {
		Write-Log "DB server connectivity issues"
	}

	foreach ($row in $table)
	{
		$platform = $row.Item("platform")
		$REQ = $row.Item("req_id")
		$Usernames = $row.Item("name")
		$Type = $row.Item("request_type_new_extn")
		$Computers = $row.Item("server_name")
		$Access = $row.Item("access_level")
		$Script:domain = $row.Item("domain")
		$Category = $row.Item("type_of_access")
		$frdate = $row.Item("start_date")
		$todate = $row.Item("end_date")
		$LastUp = $row.Item("last_updated")
		$Email = $row.Item("Requestor_Email_id")
		#Write-host "$REQ : $usernames : $type : $computers : $Access : $Dcaps : $Category"
		#break;
		if ($platform -eq "Wintel")
		{
			if ($Category -eq "Server access")
			{
				if ($domain -eq "CATE")
				{
					Server-Access
				}
			} 
			elseif ($Category -eq "Workstation access")
			{
				$Workstations = $row.Item("workstation_name")
				Workstation-Access
			} 
			elseif ($Category -eq "Application")
			{
				$Pids = $row.Item("pids")
				$groupnames = $row.Item("ad_group")
				Set-Application
			}
			elseif ($Category -eq "Control M")
			{
				$Pids = $row.Item("pids")
				Set-ControlM
			} 
			elseif ($Category -eq "Staging Path")
			{
				$Pathlist = $row.Item("path")
				$Pids = $row.Item("pids")
				$domain = "CATE"
				Set-StagingPath
			} 
			elseif ($Category -eq "Application Shared Folder Access")
			{
				$Pathlist = $row.Item("path")
				$Pids = $row.Item("pids")
				if ($domain -eq "CATE")
				{
					Set-ApplicationPath
				}
			}
			elseif ($Category -eq "Account Unlock")
			{
				$Pids = $row.Item("pids")
				Set-AccountUnlock
			} 
			elseif ($Category -eq "Password Reset")
			{
				$Pids = $row.Item("pids")
				Set-Password
			} 
			elseif ($Category -eq "Database Access")
			{
				$DBNames = $row.Item("db_name")
				$Pids = $row.Item("pids")
				$SQLusername = $row.Item("name")
				if ($domain -eq "CATE")
				{
					Set-DBAccess
				}
			} 
			elseif ($Category -eq "SQL User Creation")
			{
				$DBName = $row.Item("db_name")
				$SQLusername = $row.Item("sql_user_name")
				$Type = "SQL User"
				if ($domain -eq "CATE")
				{
					New-SQLUser
				}
			}  
		} elseif ($Platform -eq "AIX")
		{
			$aixmail = $null
			$aixdate = $row.Item("end_date")
			$todate = "{0:MMddmmssyy}" -f (get-date $aixdate)
			$AixGroup = $row.Item("ad_group")
			$Pids = $row.Item("pids")
			$Pathlist = $row.Item("path")
			$pidlist = ($Pids -split ',').trim()
			foreach ($aixpid in $pidlist)
			{
				if ($aixmail) {	$aixmail += ',' }
				$aixname = Get-ADUser -filter "Samaccountname -eq '$aixpid'"
				if ($aixname)
				{
					$aixfn = $aixname.givenname -replace '\s',''
					$aixln = $aixname.surname -replace '\s',''
					$aixto = $aixfn + '.' + $aixln + '@marks-and-spencer.com'
				} else {
					$aixto = "NA"
				}
				$aixmail += $aixto
			}
			if (($Pathlist.length -eq 1) -or ($Pathlist -eq $null) -or ($Pathlist.length -eq 0))
			{
				$Pathlist = "NA"
			}
			if (($AixGroup.length -eq 1) -or ($AixGroup -eq $null) -or ($AixGroup.length -eq 0))
			{
				$AixGroup = "NA"
			}
			if (($Pids.length -eq 1) -or ($Pids -eq $null) -or ($Pids.length -eq 0))
			{
				$Pids = "NA"
			}
			if (($Access.length -eq 1) -or ($Access -eq $null) -or ($Access.length -eq 0))
			{
				$Access = "NA"
			}
			AIX-Access
		}
	}
} Complete-ZeroTouch