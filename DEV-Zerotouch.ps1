# ##################################################################################################################################
# NAME: Dev-ZeroTouch.ps1
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

Function Write-Log([string]$message)
{
   Out-File -InputObject $message -FilePath $LogFile -Append
   #$LogFile = "C:\Windows\System32\Software\Tool\Zero-Touch\Logs\Common.log"
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
		$body += "Requested SQL account has been created in server $computers. Please find the credentials below <br><br>"
		$body += "<u>Credentials</u>: <br>"
		$body += "<b>Username</b>: $User <br>"
		$body += "<b>Password</b>: $Password <br><br>"
		$body += "Regards,<br>Non-Prod Wintel<br><br>"
		$Body += "*** Please do not reply, This is a system generated email *** <br>"
		$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
		$sub = "$REQ`:SQL account credentials"
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
		$script:Currdom = [Environment]::UserDomainName
		$script:DBlog = "AccessManagement"
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
		$net = new-object -ComObject WScript.Network
		$Driveinfo = new-object system.io.driveinfo("O:")
		if($Driveinfo.drivetype -eq "Network") {
		    $net.RemoveNetworkDrive("O:", $true)
		}
		$stgpwd = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($P_PWord))
		$net.MapNetworkDrive("O:", "\\10.128.11.10\NPEMS\SIP\UGP\Wintel\Access_Output\ZeroTouch", $false, "$mailuser", "$stgpwd") 
	} catch [System.IO.FileNotFoundException] {
		Write-Log "Active Directory Tools were not installed in this Machine, Please install this feature"
	} catch {
		Write-Log "ZeroTouch Dev Main Exception" -ForegroundColor Yellow
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
	}
}

########################################
# DB Logs Functions
########################################

Function Database-Log ([Switch]$New, [Switch]$Extension)
{
	$Usertbl = "DBAccess"
	$DBlogdb = New-Object Microsoft.SqlServer.Management.Smo.Database
	$DBlogdb = $srv.Databases.Item("$DBlog")
	$DBlogdb.ExecuteNonQuery("CHECKPOINT")
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
			$result = $DBlogdb.ExecuteWithResults($sqltext);
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
			$sqltext = "insert into [$DBlog].[dbo].[$Usertbl] values ('$Today','$Currdom','$User','$password','$Type','SUCCESS','SQL User Created','$Email','$Req');"
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
			#write-host "$sqltext"
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
		$upqry = "UPDATE [$sqldb].[dbo].[$SQLtbl] SET current_status='Completed',comments=`'$Category`: `'$cmt ,last_updated=getdate() WHERE req_id=`'$REQ`';"
		$completion = $db.ExecuteWithResults("$upqry")
		Write-Log ""
		Write-Log "Request Completed"
	} else {
		$upqry = "UPDATE [$sqldb].[dbo].[$SQLtbl] SET current_status='Completed',comments=`'Exception Occured: `'$cmt ,last_updated=getdate() WHERE req_id=`'$REQ`';"
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
		$LogFile = "O:\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\Server-Access.log"
		Write-Log -Message "Server Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
        $i = 0
		$commentlog = $null
		$commentlog = @()
        $compnames = ($computers.tostring().ToUpper() -split ',') | %{$_.Trim()}
        $ccount = $compnames.count
        $users = ($usernames -split ',') | %{$_.Trim()}
        $ucount = $users.count
        $count = $ccount * $ucount
		foreach ($computer in $compnames)
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
				$DomainGroup = [ADSI]"WinNT://MNSUKDEV/$Groupname,group"
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
						Write-Log "Error while adding domain group `"$groupname`" to local group($access). Please check with Non-Prod Wintel team"
						$commentlog += "Error while adding domain group `"$groupname`" to local group($access). Please check with Non-Prod Wintel team;"
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
		}
	} catch {
		Write-Log "Server Access Exception"
        write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
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
		$LogFile = "O:\$REQ.log"
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
			if ($path[12] -eq 'D')
			{	
				if (Test-path $path)
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
					Write-Log "Given share path not found"
					$commentlog += "Given share path not found;"
				}
			} else {
				Write-Log "Given share path is not a DEV server"
				$commentlog += "Given share path is not a DEV server;"
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

Function Set-DBAccess
{
	Start-Main
	try
	{
		$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
		$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		$LogFile = "O:\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\DBAccess.log"
		Write-Log -Message "Database Access: $Today`: $REQ ($Dcaps) - Type: $Type Request"
		Write-Log -Message "=========================================================================================="
		$i = 0
        $Users = ($SQLusername -split ',') | %{$_.Trim()}
		$DBs = ($DBNames -split ',') | %{$_.Trim()}
        $ucount = $Users.count
		$dcount= $DBs.count
		$count = $ucount * $dcount
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
		$Svr = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $Computers
		if ($Access -eq "Admin")
		{
			$Role = "db_owner"
		} elseif ($Access -eq "Regular")
		{
			$Role = "db_datareader"
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
									}
								} else {
									Write-Log "User `"$user`" not created on server `"$computers`""
									$commentlog += "User `"$user`" not created on server `"$computers`" ;"
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
		$From = "{0:dd-MMM-yy}" -f (Get-date $frdate)
		$To = "{0:dd-MMM-yy}" -f (Get-date $todate)
		$LogFile = "O:\$REQ.log"
		#$LogFile = "C:\Wintel-SIP\Zero-Touch\Logs\SQLUser-Creation.log"
		Write-Log -Message "SQL User Creation: $Today`: $REQ ($Domain)"
		Write-Log -Message "=========================================================================================="
		$i = 0
		$Users = ($SQLusername -split ',') | %{$_.Trim()}
		$count = $Users.count
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
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
					Write-log "SQL User `"$user`" added to server `"$computers`""
					$commentlog += "SQL User `"$user`" added to server `"$computers`" ;"
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
						$totchk = ++$i
						UserDB-update
						UserMgmt-SendEmail
						Write-Log "SQL Logon`:`"$User`" provided access to `"$DBName`" database as `"$Role`" on Server `"$computers`""
						Write-log "Login Credentials:"				
						Write-log "Username: $User "
						Write-Log "Password: $Password "
						$commentlog += "SQL Logon`:`"$User`" provided access to `"$DBName`" database as `"$Role`" on Server `"$computers`" ;"
						$commentlog += "Contact Non-prod team for SQL login credentials ;"
						#Write-Output $Output | Export-Csv \\10.128.11.10\npems\SIP\UGP\Wintel\Access_Output\DBAccessAuditLog.csv -NoTypeInformation -Append
					}
				}
			}
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


########################################
# Root Function
########################################

Function Complete-ZeroTouch
{
	try
	{
		$SQLSvr = "MSHSRMNSUKC0142"
		$sqldb = "ZeroTouchDB"
		$SQLtbl = "tbl_zero_touch_access"
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
		[Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMOExtended") | out-null
		$Script:srv = new-object('Microsoft.SqlServer.Management.Smo.Server') $SQLSvr
		$srv.ConnectionContext.LoginSecure = $false
		$srv.ConnectionContext.set_Login("ZeroSQL")
		$encryptedpw = "76492d1116743f0423413b16050a5345MgB8AEEAUQBhAHMAVwA5AFcAcQBQAFMAdgBtAEcAKwB6AG4AdgBuAHkAcQA5AGcAPQA9AHwAZgA1AGEAMABiADIAZQBhAGQAMwBjAGIAMwA3AGIAMQBhAGQAYwAzADkAZgA3ADUAOABhADAAOQAxADgAZAAzADEAOQBhADYAOAA3ADkANQAwADgAMwA2ADEAYQAyAGEAYQA2AGIAMAAxAGYANQBiADMAMgA3ADcAYQBiADQAZQA="
		$pwKey = (2,3,1,2,55,33,253,221,0,0,1,22,41,55,34,234,2,33,3,8,7,6,36,44)
		$SqlUPwd = $encryptedpw | ConvertTo-SecureString -Key $pwkey
		$srv.ConnectionContext.set_SecurePassword($SqlUpwd)
		$db = New-Object Microsoft.SqlServer.Management.Smo.Database
		$db = $srv.Databases.Item("$sqldb")
		$db.ExecuteNonQuery("CHECKPOINT")
		if (($srv.Information).ComputerNamePhysicalNetBIOS -match $SQLSvr)
		{
			if (($srv.Databases | ? {$_.Name -eq "$sqldb"}).Name -Match $SQLdb)
			{
				if ((($srv.Databases | ? {$_.Name -eq "$sqldb"}).tables | ? {$_.Name -eq $sqltbl}).name -match $sqltbl)
				{
					$sqltext = "Select * from [$sqldb].[dbo].[$SQLtbl] Where current_status = 'ImplementationInProgress';"
					$result = $db.ExecuteWithResults($sqltext);
					$table = $result.Tables[0];
				} else {
					Write-Log "Table `"$SQLtbl`" not found"
				}
			} else {
				Write-Log "DB `"$sqldb`" not found"
			}
		} else {
			Write-Log "DB server connectivity issues"
		}

		foreach ($row in $table)
		{
			$platform = $row.Item("platform")
			$REQ = $row.Item("req_id")
			$Type = $row.Item("request_type_new_extn")
			$Computers = $row.Item("server_name")
			$Access = $row.Item("access_level")
			$Script:domain = $row.Item("domain")
			$Category = $row.Item("type_of_access")
			$frdate = $row.Item("start_date")
			$todate = $row.Item("end_date")
			$Email = $row.Item("Requestor_Email_id")
			#$LastUp = $row.Item("last_updated")
		    #$Dcaps = $row.Item("Domain")
		    #Write-host "$REQ : $platform : $type : $computers : $Access : $domain : $Category"
			
			if ($platform -eq "Wintel")
			{
				if ($domain -eq "DEV")
				{
					if ($Category -eq "Server access")
					{
						$Usernames = $row.Item("name")
						Server-Access
					}
					elseif ($Category -eq "Application Shared Folder Access")
					{
						$pathlist = $row.Item("path")
						$Pids = $row.Item("pids")
						if ($domain -eq "DEV")
						{
							Set-ApplicationPath
						}
					}
					elseif ($Category -eq "Database Access")
					{
						$DBNames = $row.Item("db_name")
						$Pids = $row.Item("pids")
						$SQLusername = $row.Item("name")
						Set-DBAccess
					} 
					elseif ($Category -eq "SQL User Creation")
					{
						write "inside"
						$DBName = $row.Item("db_name")
						$SQLusername = $row.Item("sql_user_name")
						New-SQLUser
					}
				}
			}	
		}
	} catch {
		Write-Log "ZeroTouch DEV Exception"
		Write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		Write-Log "Exception Message: $($_.Exception.Message)"
	}
	#write-host "Hello World"
} Complete-ZeroTouch