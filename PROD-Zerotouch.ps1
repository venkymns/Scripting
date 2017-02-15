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
# 1.1 2016.05.14 Upgraded script with DB logs
#
# ##################################################################################################################################


########################################
# Common Functions
########################################

Function Write-Log([string]$message)
{
   Out-File -InputObject $message -FilePath $LogFile -Append
}

Function UserMgmt-SendEmail
{
	#Write-Host "Test"
	#Send-MailMessage -From "Non-Prod User Management <Non-ProdUserManagement@marksandspencercate.com>" -To "balaji.srinivasan@marks-and-spencer.com" -Subject "TEst" -Body "TEsting" -SmtpServer 10.151.209.17 -Credential $Mailcred -EA Stop
	$fn = $name.givenname
	$sn = $name.surname
	#$sam = $name.SamAccountName
	$doma = $serverdom.Split('.')[0].ToString().Toupper()
	$bcc = "grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com"
	#$Curruser = [Environment]::username
	if ($samid -match "\d")
	{
		$Proddata = Get-ADUser -Filter "samaccountname -eq '$samid'" -Properties UserPrincipalName
		$to = ($Proddata.UserPrincipalName).Split('@')[0]
	} else {
		$Proddata = Get-ADUser -Filter "UserPrincipalName -like '$samid*@mnscorp.net'" -Properties UserPrincipalName
		if ($Proddata)
		{
			if ($Proddata -is [System.Array])
			{
				if ($Proddata[0].userprincipalname.split('@')[1] -eq "mnscorp.net") 
				{
					$to = ($Proddata[0].UserPrincipalName).Split('@')[0]
				} elseif ($Proddata[1].userprincipalname.split('@')[1] -eq "mnscorp.net")  {
					$to = ($Proddata[1].UserPrincipalName).Split('@')[0]
				} else {
					$to = ($Proddata[0].UserPrincipalName).Split('@')[0]
				}
			} else {
				$to = ($Proddata.UserPrincipalName).Split('@')[0]
			}
		} else {
		  Write-Log "Unable to send e-mail, Please communicate with user"
		}
	}
	if ($Category -eq "Account Extension")
	{
		$body = "Hello <b>$fn</b> <br><br>"
		$body += "Your <b>$doma\$samid</b> account has been extended to: <b>$print</b> <br><br>"
        $body += "<b>Note</b>: Try logging with your user account after 2 to 3 mins<br><br>"
		$body += "Regards,<br>Non-Prod Wintel<br><br>"
		$Body += "*** Please do not reply, This is a system generated email *** <br>"
		$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
		$sub = "$REQ`:New $doma Account Expiry Date"
	} elseif ($Category -eq "User Account Creation")
	{
		if ($Type -eq 'Y-Account')
		{
			$to = ($Email -split '@')[0]
			$yfn = (($Email -split '@')[0]).Split('.')[0]
			$Body = "Hello <b>$yfn</b> <br><br> "
			$body += "New Y account has been created in $doma domain. Please find the login credential below<br><br>"
			$body += "Username: <b>$doma\$samid</b> <br>"
			$body += "Password: <b>$password</b> <br><br>"
			$body += "<b>Note</b>: Use your new account after 3 to 4 mins<br><br>"
			$body += "Regards,<br>Non-Prod Wintel<br><br>"
			$Body += "*** Please do not reply, This is a system generated email *** <br>"
			$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
			$sub = "$REQ`:New $doma Y-account"
		} else {
			$Body = "Hello <b>$fn</b> <br><br> "
			$body += "New user account has been created for you in $doma domain. Please find the login credential below<br><br>"
			$body += "Username: <b>$doma\$samid</b> <br>"
			$body += "Password: <b>$password</b> <br><br>"
			$body += "<b>Note</b>: Use your new account after 3 to 4 mins<br><br>"
			$body += "Regards,<br>Non-Prod Wintel<br><br>"
			$Body += "*** Please do not reply, This is a system generated email *** <br>"
			$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
			$sub = "$REQ`:New $doma account"
		}
	}
    
	try
	{        
    	Send-MailMessage -From "Non-Prod Access Management <Non-ProdAccessManagement@marksandspencercate.com>" -To "$to@marks-and-spencer.com" -BCC "$bcc" -Subject $sub -Body $Body -BodyAsHtml -SmtpServer 10.151.209.17 -Credential $Mailcred -EA Stop
		#Send-MailMessage -From "Non-Prod Access Management <Non-ProdAccessManagement@marksandspencercate.com>" -To "balaji.srinivasan@marks-and-spencer.com" -Subject $sub -Body $Body -BodyAsHtml -SmtpServer 10.151.209.17 -Credential $Mailcred -EA Stop
	} catch [System.Net.Mail.SmtpException] {
		Write-Log "Email not sent to user $sam`: Invalid Email Address"
	} catch {
        Write-Log "Prod Mail Exception"
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
		#$script:commentlog = $null
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
		$Driveinfo = new-object system.io.driveinfo("N:")
		if($Driveinfo.drivetype -eq "Network") {
		    $net.RemoveNetworkDrive("N:", $true)
		}
		$stgpwd = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($P_PWord))
		$net.MapNetworkDrive("N:", "\\10.128.11.10\G$\NPEMS_TEMP\NPEMS\SIP\UGP\Wintel\Access_Output\ZeroTouch", $false, "$mailuser", "$stgpwd")
	} catch [System.Management.Automation.RuntimeException] {
		Write-Log "We are unable to send password to user `"$pid`", since this account is not available in PROD"
	} catch [System.IO.FileNotFoundException] {
		Write-Log "Active Directory Tools were not installed in this Machine, Please install this feature"
	} catch {
		Write-Log "ZeroTouch PROD Main Exception"
		Write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		Write-Log "Exception Message: $($_.Exception.Message)"
	}
}

Function Write-Comment
{
	$cmt = $null
	#$commentlog
	$cmtsec = $commentlog -split ';'
	foreach ($sec in $cmtsec)
	{
		$cmt += " + Char(13) + " + "`'$sec`'"
	}
	Write-Host "$count : $totchk"
	if ($Count -eq $totchk)
	{
		$upqry = "UPDATE [ZeroTouchDB].[dbo].[tbl_zero_touch_access] SET current_status='Completed',comments=`'$Category`: `'$cmt ,last_updated=getdate() WHERE req_id=`'$REQ`';"
		$completion = $db.ExecuteWithResults("$upqry")
		Write-Log ""
		Write-Log "Request Completed"
	} else {
		$upqry = "UPDATE [ZeroTouchDB].[dbo].[tbl_zero_touch_access] SET current_status='Completed',comments=`'Exception Occured: `'$cmt ,last_updated=getdate() WHERE req_id=`'$REQ`';"
		$completion = $db.ExecuteWithResults("$upqry")
		Write-Log ""
		Write-Log "Request Not Completed"
	}#>
}

########################################
# Access Functions
########################################

Function Set-Extension
{
	Start-Main
	try 
	{
		#$LogFile = "C:\Windows\System32\Software\Tool\Zero-Touch\Logs\$REQ.log"
		$LogFile = "N:\$REQ.log"
		Write-Log -Message "Account Extension: $Today`: $REQ ($Domain)"
		Write-Log -Message "===================================================="
		$i = 0
		$commentlog = $null
		$commentlog = @()
		$Users = ($Pids -split ',') | %{$_.Trim()}
		$count = @($Users).count
		foreach ($user in $users)
		{
			$name = Get-ADUser -Filter "samaccountname -eq '$user'" -Credential $Creds -Server $serverdom -Properties LockedOut, AccountExpirationDate
			if ($name)
			{
				if ($($name.AccountExpirationDate))
				{
					if ($($name.AccountExpirationDate) -lt $(Get-Date))
					{
						if ($user -match "\d")
						{
							$Prod = Get-ADUser -Filter "samaccountname -eq '$user'" -Properties AccountExpirationDate
						} else {
							$Prod = Get-ADUser -Filter "UserPrincipalName -like '$user*@mnscorp.net'" -Properties AccountExpirationDate
						}
						if ($Prod)
						{
							if ($Prod -is [System.Array])
							{
								if ($Prod[0].userprincipalname.split('@')[1] -eq "mnscorp.net") 
								{
									$Prod = $Prod[0]
								} elseif ($Prod[1].userprincipalname.split('@')[1] -eq "mnscorp.net") {
									$Prod = $Prod[1]
								} else {
									$Prod = $Prod[0]
								}
							}
							if ($($Prod.AccountExpirationDate))
							{
								$date = $Prod.AccountExpirationDate
								$print = "{0:dd/MMM/yy}" -f $date
								Set-ADAccountExpiration -Server $serverdom -Identity "$user" -Credential $Creds -DateTime "$($date.ToShortDateString())"
								if($?)
								{
									$samid = $Name.samaccountname
									UserMgmt-SendEmail
									Write-Log "$user`: Account Extended : $Print"
									$commentlog += "$user`: Account Extended : $Print`;"
									$totchk = ++$i
								} else {
									write-Log "$user`: Unknown issue, Check with Non-Prod Wintel"
									$commentlog += "$user`: Unknown issue, Check with Non-Prod Wintel;"
								}
							} else {
								Write-Log "$user`: Prod Account Set To Never"
								$date = (get-date).addmonths(6)
								$print = "{0:dd/MMM/yy}" -f $date
								Set-ADAccountExpiration -Server $serverdom -Identity "$user" -Credential $Creds -DateTime "$($date.ToShortDateString())"
								if($?)
								{
									UserMgmt-SendEmail
									Write-Log "$user`: Account Extended to six months : $Print"
									$commentlog += "$user`: Account Extended to six months : $Print`;"
									$totchk = ++$i
								} else {
									write-Log "$user`: Unknown issue, Check with Non-Prod Wintel"
									$commentlog += "$user`: Unknown issue, Check with Non-Prod Wintel;"
								}
							}
						} else {
							Write-Log "$user`: User Not Found in Production domain"
							$commentlog += "$user`: User Not Found in Production domain;"
						}
					} else {
						Write-Log "$user`: User Account Not Expired in $domain domain"
						$commentlog += "$user`: User Account Not Expired in $domain domain;"
					}
				} else {
					Write-Log "$user`: User Account Set To Never in $domain domain"
					$commentlog += "$user`: User Account Set To Never in $domain domain;"
				}
			} else {
				Write-Log "$user`: User Not Found in $domain domain"
				$commentlog += "$user`: User Not Found in $domain domain;"
			}
		}
	} catch {
		Write-Log "Extension Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "===================================================="
}

Function New-Group
{
	Start-Main
	try
	{
		#$LogFile = "C:\Windows\System32\Software\Tool\Zero-Touch\Logs\Create-Group.log"
		$LogFile = "N:\$REQ.log"
		Write-Log -Message "Create New Group Request: $Today`: $REQ ($Dcaps)"
		Write-Log -Message "=========================================================================================="
		$i = 0
		$commentlog = $null
		$commentlog = @()
		$Users = ($Pids -split ',') | %{$_.Trim()}
		$count = @($Users).count
		if ($Currdom -eq "MNSUK")
		{
			foreach ($user in $users)
			{
				$Name = Get-ADGroup -Filter "SamAccountName -eq '$user'" -Properties GroupCategory, GroupScope, DistinguishedName
				if ($Name -ne $null)
				{
					$samid = $Name.SamAccountName
					$prechk = Get-ADGroup -Server $serverdom -Credential $creds -Filter "SAMAccountname -eq '$samid'"
					if (!($prechk))
					{
						$OUName = (((($Name.DistinguishedName -split ",", 2)[1]).split(',') |  ? {$_ -like "OU=*"}) -join ",") + ",$DC"
						$Scope = $Name.Groupscope
						$Category = $name.GroupCategory
						New-ADGroup -Server $serverdom -Credential $Creds -SamAccountName "$samid" -Path "$OUName" -Name "$samid" -GroupCategory $Category -GroupScope $Scope
						if ($?)
						{
							$chkcre = Get-ADGroup -Server $serverdom -Credential $creds -Filter "SAMAccountname -eq '$samid'"
							if ($Chkcre)
							{
								Write-Log "$user : Group Created in $domain domain"
								$commentlog += "$user : Group Created in $domain domain;"
								$totchk = ++$i
							} else {
								Write-Log "$user: Group Not Created, Please check with Non-Prod Wintel"
								$commentlog += "$user: Group Not Created, Please check with Non-Prod Wintel;"
							}
						} else {
							Write-Log "$user: Issue with creating group, please check with admin"
							$commentlog += "$user: Issue with creating group, please check with admin;"
							Write-Log "$error[0].exception.message"
							$commentlog += "$error[0].exception.message;"
						}
					} else {
						Write-Log "$user : Group Found in Non-Prod"
						$commentlog += "$user : Group Found in Non-Prod;"
					}
				} else {
					Write-Log "$user : Group Not Found in Production"
					$commentlog += "$user : Group Not Found in Production;"
				}
			}
		} else {
			Write-Log "Execute this script from Production domain"
			$commentlog += "Execute this script from Production domain;"
		}
	} catch {
		Write-Log "Extension Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	Write-Comment
	Write-Log -Message "=========================================================================================="
}

Function New-User
{
	Start-Main
	$commentlog = $null
	$commentlog = @()
	Function UserDB-update
	{
		$CreateDB = "AccessManagement"
		$Usertbl = "UserGroup-Creation"
		$Userdb = New-Object Microsoft.SqlServer.Management.Smo.Database
		$userdb = $MySQL.Databases.Item("$CreateDB")
		$userdb.ExecuteNonQuery("CHECKPOINT")
		if (($MySQL.Databases | ? {$_.Name -eq "$CreateDB"}).Name -Match $CreateDB)
		{
			if ((($MySQL.Databases | ? {$_.Name -eq "$CreateDB"}).tables | ? {$_.Name -eq $Usertbl}).name -match $Usertbl)
			{
				$sqltext = "insert into [$CreateDB].[dbo].[$Usertbl] values ('$Today','$($serverdom.Split('.')[0])','$Samid','$password','$Type','SUCCESS','User Account Created','$Email','$REQ');"
				$result = $userdb.ExecuteWithResults($sqltext);
			} else {
				Write-Log "Table `"$Usertbl`" not found"
			}
		} else {
			Write-Log "DB `"$CreateDB`" not found"
		}
	}
	<#Function Create-User ([Switch]$PIDUser, [Switch]$Nameuser, [Switch]$YAccuser)
	{
		if ($PIDUser)
		{
			$Display = $Name.SamAccountName
			$CPAL = $False
			$CCP = $False
			$PNE = $False
			$userPname = $samid + "@$serverdom"
		} elseif ($Nameuser)
		{
			$Display = $UPN
			$CPAL = $True
			$CCP = $False
			$PNE = $False
			$userPname = $samid + "@$serverdom"
		} elseif ($YAccuser)
		{
			$Display = $Name.SamAccountName
			$CPAL = $False
			$CCP = $True
			$PNE = $True
		}
		New-ADUser -Server $serverdom -Credential $Creds -SamAccountName "$samid" -Path "$OUName" -GivenName $Name.GivenName `
		-Surname $Name.Surname -DisplayName $Display -Name "$Samid" -Office $Name.Office -AccountExpirationDate $dat -Description $Name.Description `
		-Title $Name.Title -Department $Name.Department -Company $Name.Company -UserPrincipalName $userPname `
		-AccountPassword (ConvertTo-SecureString -AsPlainText "$password" -Force) -ChangePasswordAtLogon $CPAL -CannotChangePassword $CCP `
		-PasswordNeverExpires $PNE -enable $true
		if ($?)
		{
			$chkcre = Get-ADUser -Server $serverdom -Credential $creds -Filter "SAMAccountname -eq '$samid'"
			if ($Chkcre)
			{
				UserDB-Update
				Write-Log "$Samid`: User Account Created in $domain domain, and login credential is : $password"
				$commentlog += "$Samid`: User Account Created in $domain domain, and login credential will be shared to user in separate mail;"
				$totchk = ++$i
			} else {
				Write-Log "$Samid`: User Not Created, Please check with admin"
				$commentlog += "$Samid`: User Not Created, Please check with non-prod team;"
			}
		} else {
			Write-Log "$Samid`: Issue with creating user, please check with admin"
			$commentlog += "$Samid`: Issue with creating user, please check with non-prod team;"
		}
	}#>
	try
	{
		if ($Currdom -eq "MNSUK")
		{
			#$LogFile = "C:\Windows\System32\Software\Tool\Zero-Touch\Logs\Create-User.log"
			$LogFile = "N:\$REQ.log"
			Write-Log -Message "Create New User Request: $Today`: $REQ ($Domain)"
			Write-Log -Message "===================================================="
			$i = 0
			$Users = ($Pids -split ',') | %{$_.Trim()}
			$count = @($Users).count
			foreach ($user in $users)
			{
				$arr = "london","kingdom","default","welcome","wisdom","happy","hello","secret","master","private","forever","spencer","destiny","summer"
				$Ranstr = Get-Random $arr
				$Rannum = Get-Random -Minimum 10 -Maximum 100
				$script:password = $Ranstr + $Rannum
				if ($Type -ne "Test Account")
				{
					$Name = Get-ADUser -Filter "SamAccountName -eq '$user'" -Properties AccountExpirationDate, Department, Description, DisplayName, Company, GivenName, Office, SamAccountName, Name, Surname, Title, DistinguishedName, UserPrincipalName
					if ($Name -ne $null)
					{
						$OUName = (((($Name.DistinguishedName -split ",", 2)[1]).split(',') |  ? {$_ -like "OU=*"}) -join ",") + ",$DC"
						$script:UPN = ($Name.UserPrincipalName).split('@')[0]
						if ($Name.AccountExpirationDate)
						{
							$dat = $(($Name.AccountExpirationDate).ToShortDateString())
						} else {
							$dat = $null
						}
						if ($Type -eq "PID")
						{
							if ($user -match "^([A-S])\d")
							{
								$samid = $Name.SamAccountName
								$pidchk = Get-ADUser -Filter "samaccountname -eq '$samid'" -Server $serverdom -Credential $Creds
								if (!($pidchk))
								{
									$Display = $Name.SamAccountName
									$CPAL = $False
									$CCP = $False
									$PNE = $False
									$userPname = $samid + "@$serverdom"
									New-ADUser -Server $serverdom -Credential $Creds -SamAccountName "$samid" -Path "$OUName" -GivenName $Name.GivenName `
									-Surname $Name.Surname -DisplayName $Display -Name "$Samid" -Office $Name.Office -AccountExpirationDate $dat -Description $Name.Description `
									-Title $Name.Title -Department $Name.Department -Company $Name.Company -UserPrincipalName $userPname `
									-AccountPassword (ConvertTo-SecureString -AsPlainText "$password" -Force) -ChangePasswordAtLogon $CPAL -CannotChangePassword $CCP `
									-PasswordNeverExpires $PNE -enable $true
									if ($?)
									{
										$chkcre = Get-ADUser -Server $serverdom -Credential $creds -Filter "SAMAccountname -eq '$samid'"
										if ($Chkcre)
										{
											UserMgmt-SendEmail
											UserDB-Update
											Write-Log "$Samid`: User Account Created in $domain domain, and login credential is : $password"
											$commentlog += "$Samid`: User Account Created in $domain domain, and login credential will be shared to user in separate mail;"
											$totchk = ++$i
										} else {
											Write-Log "$Samid`: User Not Created, Please check with admin"
											$commentlog += "$Samid`: User Not Created, Please check with non-prod team;"
										}
									} else {
										Write-Log "$Samid`: Issue with creating user, please check with admin"
										$commentlog += "$Samid`: Issue with creating user, please check with non-prod team;"
									}
									#Create-User -PIDUser
								} else {
									write-log "$user`: This PID found in Non-Prod $d domain"
									$commentlog += "$user`: This PID found in Non-Prod $d domain;" 
								}
							} else {
								write-log "$user`: This is not a PID user"
								$commentlog += "$user`: This is not a PID user;"
							}
						}
						elseif ($Type -eq "Name.Name")
						{
							if ($user -match "^([A-S])\d")
							{
								if ($UPN.Length -gt "20")
								{
									$samid = $UPN.Substring(0,20)
								} else {
									$samid = $UPN
								}
								$namechk = Get-ADUser -Filter "samaccountname -eq '$samid'" -Server $serverdom -Credential $Creds
								if (!($namechk))
								{
									$Display = $UPN
									$CPAL = $True
									$CCP = $False
									$PNE = $False
									$userPname = $samid + "@$serverdom"
									New-ADUser -Server $serverdom -Credential $Creds -SamAccountName "$samid" -Path "$OUName" -GivenName $Name.GivenName `
									-Surname $Name.Surname -DisplayName $Display -Name "$Samid" -Office $Name.Office -AccountExpirationDate $dat -Description $Name.Description `
									-Title $Name.Title -Department $Name.Department -Company $Name.Company -UserPrincipalName $userPname `
									-AccountPassword (ConvertTo-SecureString -AsPlainText "$password" -Force) -ChangePasswordAtLogon $CPAL -CannotChangePassword $CCP `
									-PasswordNeverExpires $PNE -enable $true
									if ($?)
									{
										$chkcre = Get-ADUser -Server $serverdom -Credential $creds -Filter "SAMAccountname -eq '$samid'"
										if ($Chkcre)
										{
											UserMgmt-SendEmail
											UserDB-Update
											Write-Log "$Samid`: User Account Created in $domain domain, and login credential is : $password"
											$commentlog += "$Samid`: User Account Created in $domain domain, and login credential will be shared to user in separate mail;"
											$totchk = ++$i
										} else {
											Write-Log "$Samid`: User Not Created, Please check with admin"
											$commentlog += "$Samid`: User Not Created, Please check with non-prod team;"
										}
									} else {
										Write-Log "$Samid`: Issue with creating user, please check with admin"
										$commentlog += "$Samid`: Issue with creating user, please check with non-prod team;"
									}
									#Create-User -Nameuser
								} else {
									write-log "$Samid`: This Name.Name found in Non-Prod $d domain"
									$commentlog += "$Samid`: This Name.Name found in Non-Prod $d domain;"
								}
							} else {
								write-log "$user`: This is not a PID user"
								$commentlog += "$user`: This is not a PID user;"
							}
						} 
						elseif ($Type -eq "Y-Account")
						{
							if ($user -match "Y\d")
							{
								$samid = $Name.SamAccountName
								$ychk = Get-ADUser -Filter "samaccountname -eq '$samid'" -Server $serverdom -Credential $Creds
								if (!($ychk))
								{
									$Display = $Name.SamAccountName
									$CPAL = $False
									$CCP = $True
									$PNE = $True
									New-ADUser -Server $serverdom -Credential $Creds -SamAccountName "$samid" -Path "$OUName" -GivenName $Name.GivenName `
									-Surname $Name.Surname -DisplayName $Display -Name "$Samid" -Office $Name.Office -AccountExpirationDate $dat -Description $Name.Description `
									-Title $Name.Title -Department $Name.Department -Company $Name.Company -UserPrincipalName $userPname `
									-AccountPassword (ConvertTo-SecureString -AsPlainText "$password" -Force) -ChangePasswordAtLogon $CPAL -CannotChangePassword $CCP `
									-PasswordNeverExpires $PNE -enable $true
									if ($?)
									{
										$chkcre = Get-ADUser -Server $serverdom -Credential $creds -Filter "SAMAccountname -eq '$samid'"
										if ($Chkcre)
										{
											UserMgmt-SendEmail
											UserDB-Update
											Write-Log "$Samid`: Y-Account Created in $domain domain, and login credential is : $password"
											$commentlog += "$Samid`: Y-Account Created in $domain domain, please contact non-prod wintel team for password;"
											$totchk = ++$i
										} else {
											Write-Log "$Samid`: User Not Created, Please check with admin"
											$commentlog += "$Samid`: User Not Created, Please check with non-prod team;"
										}
									} else {
										Write-Log "$Samid`: Issue with creating user, please check with admin"
										$commentlog += "$Samid`: Issue with creating user, please check with non-prod team;"
									}
									#Create-User -YAccuser
								} else {
									write-log "$samid`: This Y account found in Non-Prod $d domain"
									$commentlog += "$samid`: This Y account found in Non-Prod $d domain;" 
								}
							} else {
								write-log "$user`: This is not a Y-Account"
								$commentlog += "$user`: This is not a Y-Account;"
							}
						}
					} else {
						Write-Log "$user`: User Not Found in Production"
						$commentlog += "$user`: User Not Found in Production;"
					}
				} else {
					$samid = $user
					$OUName = "OU=HO Users,OU=User Accounts,DC=mnsuk$d,DC=adroot$d,DC=marksandspencer$d,DC=com"
					$dat = (get-date).AddMonths(6).ToShortDateString()
					$userPname = $samid + "@$serverdom"
					$CPAL = $False
					$CCP = $False
					$PNE = $False
					$tchk = Get-ADUser -Filter "samaccountname -eq '$samid'" -Server $serverdom -Credential $Creds
					if (!($tchk))
					{
						New-ADUser -Server $serverdom -Credential $Creds -SamAccountName "$samid" -Path "$OUName" -GivenName "$samid" `
						-DisplayName "$Samid" -Name "$Samid" -AccountExpirationDate $dat -UserPrincipalName $userPname -ChangePasswordAtLogon $CPAL `
						-AccountPassword (ConvertTo-SecureString -AsPlainText "$password" -Force) -CannotChangePassword $CCP `
						-PasswordNeverExpires $PNE -enable $true
						if ($?)
						{
							$chkcre = Get-ADUser -Server $serverdom -Credential $creds -Filter "SAMAccountname -eq '$samid'"
							if ($Chkcre)
							{
								UserDB-Update
								Write-Log "$user : Test Account Created in $domain domain, and login credential is : $password"
								$commentlog += "$user : Test Account Created in $domain domain, Please contact non-prod wintel team for password;"
								$totchk = ++$i
							} else {
								Write-Log "$user: User Not Created, Please check with admin"
								$commentlog += "$user: User Not Created, Please check with admin;"
							}
						} else {
							Write-Log "$user: Issue with creating user, please check with admin"
							$commentlog += "$user: Issue with creating user, please check with admin;"
						}
					} else {
						write-log "$samid`: This T account found in Non-Prod $d domain"
						$commentlog += "$samid`: This T account found in Non-Prod $d domain;" 
					}
				}
			}
		} else {
			Write-Log "Execute this script from Production domain"
			$commentlog += "Execute this script from Production domain;"
		}
	} catch {
		Write-Log "User Creation Exception"
		write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		write-Log "Exception Message: $($_.Exception.Message)"
		$commentlog += "Exception Message: $($_.Exception.Message);"
	}
	<#if (($Type -eq "Test Account") -or ($Type -eq "Y-Account"))
	{
		if ($Count -eq $totchk)
		{
			$Log = "N:\$REQ.log"
			$bcc = "grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com"
			$MailName = $Email.Split('.')[0]
			$Body = "Hello <b>$MailName</b> <br><br> "
			$body += "Please find the attached login credentials for the given accounts<br><br>"
			$body += "<b>Note</b>: Try logging with account after 3 to 4 mins<br><br>"
			$body += "Regards,<br>Non-Prod Wintel<br><br>"
			$Body += "*** Please do not reply, This is a system generated email *** <br>"
			$Body += "*** Please contact `"grp-dl-it-tcs-nonprodwintelsupport@mnscorp.onmicrosoft.com`" for any issues ***"
			$sub = "Login credentials for $doma domain"
			Send-MailMessage -From "Non-Prod User Management <Non-ProdUserManagement@marksandspencercate.com>" -To "$Email@marks-and-spencer.com" -BCC "$bcc" -Subject $sub -Attachments $Log -Body $Body -BodyAsHtml -SmtpServer 10.151.209.17 -Credential $Mailcred -EA Stop
		}
	}#>
	Write-Comment
	Write-Log -Message "===================================================="
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
		$Script:MySQL = new-object('Microsoft.SqlServer.Management.Smo.Server') $SQLSvr
		$MySQL.ConnectionContext.LoginSecure = $false
		$MySQL.ConnectionContext.set_Login("ZeroSQL")
		$encryptedpw = "76492d1116743f0423413b16050a5345MgB8AEEAUQBhAHMAVwA5AFcAcQBQAFMAdgBtAEcAKwB6AG4AdgBuAHkAcQA5AGcAPQA9AHwAZgA1AGEAMABiADIAZQBhAGQAMwBjAGIAMwA3AGIAMQBhAGQAYwAzADkAZgA3ADUAOABhADAAOQAxADgAZAAzADEAOQBhADYAOAA3ADkANQAwADgAMwA2ADEAYQAyAGEAYQA2AGIAMAAxAGYANQBiADMAMgA3ADcAYQBiADQAZQA="
		$pwKey = (2,3,1,2,55,33,253,221,0,0,1,22,41,55,34,234,2,33,3,8,7,6,36,44)
		$SqlUPwd = $encryptedpw | ConvertTo-SecureString -Key $pwkey
		$MySQL.ConnectionContext.set_SecurePassword($SqlUpwd)
		$db = New-Object Microsoft.SqlServer.Management.Smo.Database
		$db = $MySQL.Databases.Item("$sqldb")
		$db.ExecuteNonQuery("CHECKPOINT")
		if (($MySQL.Information).ComputerNamePhysicalNetBIOS -match $SQLSvr)
		{
			if (($MySQL.Databases | ? {$_.Name -eq "$sqldb"}).Name -Match $SQLdb)
			{
				if ((($MySQL.Databases | ? {$_.Name -eq "$sqldb"}).tables | ? {$_.Name -eq $sqltbl}).name -match $sqltbl)
				{
					$sqltext = "Select * from [ZeroTouchDB].[dbo].[tbl_zero_touch_access] Where current_status = 'ImplementationInProgress';"
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
			$Usernames = $row.Item("name")
			$Type = $row.Item("request_type_new_extn")
			$Computers = $row.Item("server_name")
			$Access = $row.Item("access_level")
			$Script:domain = $row.Item("domain")
			$Category = $row.Item("type_of_access")
			$From = $row.Item("start_date")
			$To = $row.Item("end_date")
			$LastUp = $row.Item("last_updated")
			$Workstations = $row.Item("workstation_name")
			$Email = $row.Item("Requestor_Email_id")
		    #$Dcaps = $row.Item("Domain")
		    #Write-log "$REQ : $usernames : $type : $computers : $Access : $domain : $Category"
			
			if ($platform -eq "Wintel")
			{
				if ($Category -eq "Account Extension")
				{
					$Pids = $row.Item("pids")
					Set-Extension
				} elseif ($Category -eq "AD Group Creation")
				{
					$Pids = $row.Item("ad_group")
					New-Group
				} elseif ($Category -eq "User Account Creation")
				{
					$Type = $row.Item("account_type")
					$Pids = $row.Item("pids")
					New-User
				}
			}	
		}
	} catch {
		Write-Log "ZeroTouch Prod Exception"
		Write-Log "Exception Type: $($_.Exception.GetType().FullName)"
		Write-Log "Exception Message: $($_.Exception.Message)"
	}
} Complete-ZeroTouch