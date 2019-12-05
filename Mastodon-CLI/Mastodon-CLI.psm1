<# 

This module frontend holds all the functions for the custom Mastodon CLI for Contoso.

    Built by [OMITTED] | 7/9/2019 | 

Version 1.1.0

#>

#Requires -RunAsAdministrator

# Set AD variables

$Global:DNSROOT = '@' + (Get-ADDomain).dnsroot
$Global:domain = (Get-ADDomain).forest
$Global:DOMAINName = (Get-ADDomain).name
 
#---------------------------------------------#

# Change me to define default settings.

$Global:DefPWD="Qwerty1"
$Global:MXIEDefPWD="74667466"
$Global:MXIEDefPIN="7466"
$Global:ExchangeServerURL = "http://wmail.olympus.contosoinc.com"
$Global:SendingEmail = "techsupport@contosoinc.com"
$Global:SMTPServer = "10.0.99.215"
$Global:MXIEURL="mxie.contosoinc.com"

#---------------------------------------------#

# Set required variables for script

$Global:StartChar = 0
$Global:EndChar = 1
$Global:ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

#---------------------------------------------#

# Script functions for miscellaneous features #

#---------------------------------------------#

Function Start-Countdown{ 

    Param(
        [Int32]$Seconds = 10,
        [string]$Message = "Pausing for 10 seconds... "
    )
    ForEach ($Count in (1..$Seconds))
    {   Write-Progress -Id 1 -Activity $Message -Status "Please standby for $Seconds seconds..., $($Seconds - $Count) left" -PercentComplete (($Count / $Seconds) * 100)
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
}

#---------------------------------------------#

# Script functions for the main CLI interface #

#---------------------------------------------#

Function Start-Mastodon{

Get-ScriptDirectory
CLS

$mastodonSPLASH = Get-Content -Path "$scriptDir\splashscreens\mastodon.splash" | Write-Host -ForegroundColor Red
		$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
CLS
Start-CLI
}

Function Start-CLI{

$startSPLASH = Get-Content -Path "$scriptDir\splashscreens\startpage.splash" | Write-Host -ForegroundColor Cyan        

            
    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host

	switch($Selection){

		1{
			CLS
			Start-ADME
			}
			
		2{
                CLS
                Start-ACME
			}
	
		3{
			 Write-Host "This feature is not available yet! Press any key to try again" -NoNewLine -ForeGroundColor Yellow
                    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL

			}

        4{
			[Console]::ForegroundColor = "Red" 
			CLS
			$GPLv3Lic = Get-Content -Path "$scriptDir\Licenses\GPLv3.lic" | Out-Host -Paging
					$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
                                  
			[Console]::ForegroundColor = "White" 
							CLS
							Start-CLI
		     }

		default{
					Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
						$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
						CLS
						Start-CLI
			}
		}
	}

Function Start-ADME{

$admeSPLASH = Get-Content -Path "$scriptDir\splashscreens\adme.splash" | Write-Host -ForegroundColor Cyan
            
    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host
		
	switch($Selection){
		1{
			CLS
			Start-UAW
			}
		
		2{
            CLS
			Start-PASSMAN
			}
				
		3{
			CLS
			Start-GRPMAN
			}
					
		4{
			CLS
			Start-CheckAD
			}		
		5{
			CLS
			Start-RemoveStaleADComputers
			}
		6{
			CLS
			Start-MoveStaleADUsers
			}        
		7{
			
			$DefaultInfo = new-object psobject -Property @{
				'DomainForest' = "$Global:domain"
				'DomainName' = $Global:DOMAINName
				'ExchangeServerURL' = $Global:ExchangeServerURL
				'ExchangeIPAddress' = $Global:SMTPServer
				'CommunicationEmail' = $Global:SendingEmail
				'HomeDrivePath' = $Global:homedir
				'MXIEURL' = $Global:MXIEURL
				}

			$DefaultInfo | Format-Table -AutoSize -Property DomainForest,DomainName,ExchangeServerURL,ExchangeIPAddress
				write-host ""
					$DefaultInfo | Format-Table -AutoSize -Property CommunicationEmail,HomeDrivePath,MXIEURL
						$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
			CLS
			Start-ADME
			}
    			           
		8{
			CLS
			Start-CLI
			}	   

		default {
			Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
				$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
			CLS
			Start-ADME
			}
	}
}

Function Start-ACME{

$acmeSPLASH = Get-Content -Path "$scriptDir\splashscreens\acme.splash" | Write-Host -ForegroundColor Cyan
            
    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host
		
	switch($Selection){
		1{
			CLS
			Start-WACHandshake
			}
		
		2{
                CLS
			    Start-DisableSMBv1
			}
				
		3{
			 Write-Host "This feature is not available yet! Press any key to try again" -NoNewLine -ForeGroundColor Yellow
                    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
                CLS
			    Start-ACME
			}
					
		4{
                CLS
			    Start-HCCM
			}

		5{
			CLS
			Start-CLI
			}
		
        default{
					Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
						$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
                CLS
			    Start-ACME
			}	

    }
}

Function Start-UAW{

$uawSPLASH = Get-Content -Path "$scriptDir\splashscreens\uaw.splash" | Write-Host -ForegroundColor Cyan
            
    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host
	switch($Selection){		
					
		1{
            CLS
            Start-Makecontoso
			}
				
        2{
            CLS
            Start-MakeWidgets
        }

        3{
            CLS
            Start-UMS
        }

        4{
            CLS
            Start-ADME
        }
        
		default {
            Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
                $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
			CLS
			Start-UAW
		}
    }
}

Function Start-PASSMAN{

	$passmanSPLASH = Get-Content -Path $scriptDir\Mastodon-CLI\Mastodon-CLI\splashscreens\passman.splash | Write-Host -ForegroundColor Cyan
            
    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host
	switch($Selection){
		1{
			CLS
			Start-UAW
			}
		
		default {
            Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
                $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
			CLS
			Start-UAW
		}
	}
}

Function Start-GRPMAN{

$grpmanSPLASH = Get-Content -Path $scriptDir\Mastodon-CLI\Mastodon-CLI\splashscreens\grpman.splash | Write-Host -ForegroundColor Cyan
            
    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host
	switch($Selection){		
		1{
			CLS
			Start-UAW
			}
		
		default {
            Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
                $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
			CLS
			Start-UAW
		}
	}
}

Function Start-UMS{

	$umsSPLASH = Get-Content -Path "$scriptDir\splashscreens\ums.splash" | Write-Host -ForegroundColor Cyan
            
    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host
	switch($Selection){		
				1{
			Write-Host "Press Any Key To Select Windows User" -ForeGroundColor Yellow
					$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
						$WinID = Get-ADUser -Filter * | Select-Object -Property SamAccountName,Name | Out-GridView -PassThru 
					
					# Seperate into Variables for Text ouptut.

					$UserSAM = $WINID | Select-Object -ExpandProperty SAMAccountName
					$UserFirst = $WINID | Select-Object -ExpandProperty Name

					# Gather new information

					Write-Host "Enter First Name: " -NoNewLine -ForeGroundColor Yellow
						$FirstName = Read-Host
							if($FirstName -eq ''){
								while($FirstName -eq ''){
								Write-Host "Invalid Entry. Try again: " -NoNewLine -ForeGroundColor Yellow
								$FirstName = Read-Host
							}
						}
								
					Write-Host "Enter Last Name: " -NoNewLine -ForeGroundColor Yellow
						$LastName = Read-Host
							if($LaseName -eq ''){
								while($LastName -eq ''){
									Write-Host "Invalid Entry. Try again: " -NoNewline -ForeGroundColor Yellow
									$LastName = Read-Host
							}
						}
					
                    # Set new information
					
					Get-ADUser -Identity "$UserSAM" | Set-ADUser #WAITING

					# Confirm information and return to menu
					
					Write-Host "$WinID is now $FirstName $LastName. Press any key to return to the menu." -ForeGroundColor Yellow
							$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
						CLS
							Start-UMS
						}
				2{
				Write-Host "Press Any Key To Select Windows User" -ForeGroundColor Yellow
						$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
							$WinID = Get-ADUser -Filter * | Select-Object -Property SamAccountName,Name | Out-GridView -PassThru 
					
						# Seperate into Variables for Text ouptut.

						$UserSAM = $WINID | Select-Object -ExpandProperty SAMAccountName
						$UserFirst = $WINID | Select-Object -ExpandProperty Name

						# Gather new information

						Write-Host "Enter Department: " -NoNewLine -ForeGroundColor Yellow
							Read-Host $Department
								if($Department -eq $null){
									while($Department -eq $NULL){
										Write-Host "Invalid Entry. Try again: " -NoNewLine -ForeGroundColor Yellow
										Read-Host $Department
								}
							}
						# Set new information
					
						Get-ADUser -Identity "$UserSAM" | Set-ADUser #WAITING

						# Confirm information and return to menu
					
						Write-Host "$WinID is now assigned to $DEPARTMENT. Press any key to return to the menu." -NoNewLine -ForeGroundColor Yellow
								$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
							CLS
								Start-UMS
				}
    					3{
							
							Write-Host "Press Any Key To Select Windows User" -NoNewLine -ForeGroundColor Yellow
								$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
								$WinID = Get-ADUser -Filter * | Select-Object -Property SamAccountName,Name | Out-GridView -PassThru 
					
								# Seperate into Variables for Text ouptut.

								$UserSAM = $WINID | Select-Object -ExpandProperty SAMAccountName
								$UserFirst = $WINID | Select-Object -ExpandProperty Name

								# Gather new information		
							
								Write-Host "Press Any Key To Select Manager" -ForeGroundColor Yellow
									$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
										$MANAGER = Get-ADUser -Filter * | Select-Object -Property SamAccountName,Name | Out-GridView -PassThru | Select-Object -ExpandProperty SAMAccountName  
										
								# Set new information
					
								Get-ADUser -Identity "$UserSAM" | Set-ADUser #WAITING
											
								# Confirm information and return to menu

								Write-Host "$UserSAM is now assigned to $MANAGER. Press any key to return to the menu." -NoNewLine -ForeGroundColor Yellow
									$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
											CLS
												Start-UMS
						}	
							4{
							Write-Host "Press Any Key To Select Windows User" -ForeGroundColor Yellow
								$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
								$WinID = Get-ADUser -Filter * | Select-Object -Property SamAccountName,Name | Out-GridView -PassThru 
					
								# Seperate into Variables for Text ouptut.

								$UserSAM = $WINID | Select-Object -ExpandProperty SAMAccountName
								$UserFirst = $WINID | Select-Object -ExpandProperty Name

								# Gather new information		
							
								Write-Host "Enter Email Address: " -NoNewLine -ForeGroundColor Yellow
									$Email = Read-Host
										if($Email -eq $null){
										while($Email -eq $NULL){
											Write-Host "Invalid Entry. Try again: " -NoNewLine -ForeGroundColor Yellow
											Read-Host $Email
											}
										}
								# Set new information
					
								Get-ADUser -Identity "$UserSAM" | Set-ADUser #WAITING
											
								# Confirm information and return to menu

								Write-Host "$UserSAM email has been updated to $Email. Press any key to return to the menu." -NoNewLine -ForeGroundColor Yellow
									$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
											CLS
												Start-UMS
								}
										5{
											CLS
											Start-UAW
												}

									default{	 
										Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
																$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
																CLS
																				Start-UMS
																				}
		}
	}

Function Start-HCCM{

$hccmSPLASH = Get-Content -Path "$scriptDir\splashscreens\hccm.splash" | Write-Host -ForegroundColor Cyan

    Write-Host "Please make a selection: " -NoNewLine -ForeGroundColor Yellow
            $Selection = Read-Host
		
	switch($Selection){
		1{
			CLS
			Start-CheckCluster
			}
		
		2{
                CLS
                Start-ExtendClusterVolume
			}
				
		3{
                CLS
			    Start-ACME
			}
        default{
			 Write-Host "Invalid selection. Press any key to try again" -NoNewLine -ForeGroundColor Yellow
				    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
                CLS
		        Start-HCCM
    
		}
	}
}

#------------------------------------------#

# Script functions for creating new users  #

#------------------------------------------#

Function Start-Makecontoso{

    # Display splash screen

    $contosoUserSPLASH = Get-Content -Path "$scriptDir\splashscreens\contosoUser.splash" | Write-Host -ForegroundColor Cyan

    # Gather user information

    Get-UserInfo

    # Set aditional user parameters

    $Global:SAM = $FirstName + $LastName.SubString($StartChar,$EndChar)
	$Global:DISPLAYNAME = "$FirstName $LastName"
	$Global:EmailName = "$FirstName.$LastName"

    Start-IfExist

    # Set additional user variables

    $Global:EmailName = "$FirstName.$LastName"
	$Global:UPN = $Global:SAM + “$Global:DNSROOT”
	$Global:DISPLAYNAME = "$LastName, $FirstName"
	$Global:EmailName = "$FirstName.$LastName"
    $Global:IsWidgets = "No"
    $Global:COMPANY = "Contoso, Inc"
    $Global:EmailDomain = "contosoinc.com"
    $Global:EmployeeEmail = "$EmailName@$EmailDomain"
    $Global:UPN = $Global:SAM + “$Global:DNSROOT”
    $ISResponse = "n"


    Start-ValidateInformation

    Start-CreateUser

    Start-CreateSecondary

    Start-CreateDrive

	Start-MakeMailbox

    Start-SendCredentials

    Start-ChangePassword

}

Function Start-MakeWidgets{

    # Display splash screen

    $ISUserSPLASH = Get-Content -Path "$scriptDir\splashscreens\ISUser.splash" | Write-Host -ForegroundColor Cyan

    # Gather user information

    Get-UserInfo

    # Set aditional variables with Get-UserInfo

    $Global:SAM = $FirstName + $LastName.SubString($StartChar,$EndChar) + "-IS"
    $Global:DISPLAYNAME = "$FirstName $LastName"
	$Global:EmailName = "$FirstName.$LastName"

    Start-IfExist

    # Set additional user variables with Check-IfExist

    $Global:EmailName = "$FirstName.$LastName"
	$Global:UPN = $Global:SAM + “$Global:DNSROOT”
	$Global:DISPLAYNAME = "$LastName, $FirstName"
	$Global:EmailName = "$FirstName.$LastName"
    $Global:IsWidgets = "Yes"
    $Global:COMPANY = "Widgets"
    $Global:DISPLAYNAME = "$DISPLAYNAME (Widgets)"
    $Global:EmailDomain = "Widgets.com"
    $Global:EmployeeEmail = "$EmailName@$EmailDomain"
    $Global:UPN = $Global:SAM + “$Global:DNSROOT”
    $ISResponse = "y"

    Start-ValidateInformation

    Start-CreateUser

    Start-CreateSecondary

    Start-CreateDrive

	Start-MakeMailbox

    Start-SendCredentials

    Start-ChangePassword
}

Function Get-UserInfo{  

  
	# Gather information


		Write-Host "`nFirst Name: " -NoNewLine -ForeGroundColor Yellow
			$Global:FirstName = Read-Host
					
			while($Global:FirstName -eq ''){
				Write-Host "Invalid Entry. First Name: " -NoNewLine -ForeGroundColor Yellow
				$Global:FirstName = Read-Host
		}

		Write-Host "Last Name: " -NoNewLine -ForeGroundColor Yellow
			$Global:LastName = Read-Host
			
			while($Global:LastName -eq ''){
				Write-Host "Invalid Entry. Last Name: " -NoNewLine -ForeGroundColor Yellow
				$Global:LastName = Read-Host
		}

		Write-Host "Department: " -NoNewLine -ForeGroundColor Yellow
			$Global:DEPARTMENT = Read-Host
			
			while($Global:DEPARTMENT -eq ''){
				Write-Host "Invalid Entry. Department: " -NoNewLine -ForeGroundColor Yellow
				$Global:DEPARTMENT = Read-Host
		}

		Write-Host "Job Title: " -NoNewLine -ForeGroundColor Yellow
			$Global:JOBTITLE = Read-Host
			
			while($Global:JOBTITLE -eq ''){
				Write-Host "Invalid Entry. Job Title: " -NoNewLine -ForeGroundColor Yellow
				$Global:JOBTITLE = Read-Host
		}

		Write-Host "MXIE Extension: " -NoNewLine -ForeGroundColor Yellow
			$Global:MXIEEXT = Read-Host
				
			while($Global:MXIEEXT -eq ''){
				Write-Host "Invalid Entry. MXIE Extension: " -NoNewLine -ForeGroundColor Yellow
				$Global:MXIEEXT = Read-Host
		}

		Write-Host "Website Username: " -NoNewLine -ForeGroundColor Yellow
			$Global:WEBACCESS = Read-Host
				
			while($Global:WEBACCESS -eq ''){
				Write-Host "Invalid Entry. Website Username: " -NoNewLine -ForeGroundColor Yellow
				$Global:WEBACCESS = Read-Host
		}

		Write-Host "Branch Number: " -NoNewLine -ForeGroundColor Yellow
			$Global:BRANCH = Read-Host
				
			while($Global:BRANCH -eq ''){
				Write-Host "Invalid Entry. Branch Number: " -NoNewLine -ForeGroundColor Yellow
				$Global:BRANCH = Read-Host
		}
		
		Write-Host "Select Manager `n" -NoNewLine -ForeGroundColor Yellow
			
			Start-Countdown -Seconds 2 -Message "Querying AD... "       
				$Global:Mgr = Get-ADUser -Filter * | Select-Object -Property SamAccountName,GivenName,Name | Out-GridView -PassThru 
                $Global:MANAGER = $Global:Mgr | Select-Object -ExpandProperty SAMAccountName
                $Global:ManagerFirst = $Global:Mgr | Select-Object -ExpandProperty GivenName
                $Global:ManagerFull = $Global:Mgr | Select-Object -ExpandProperty Name

		# Set Manager & OU container to store user object and create account
    
	Write-Host "Select AD Container`n" -nonewline -foregroundcolor Yellow 

    Start-Countdown -Seconds 2 -Message "Querying OU containers... "
		$Global:OU = Get-ADOrganizationalUnit -Filter * | Select-Object -Property DistinguishedName | Out-GridView -PassThru | Select-Object -ExpandProperty DistinguishedName 

    Write-Host "`nInformation For $FirstName $LastName collected..." -foregroundcolor Green

}

Function Start-IfExist{

	# Check if AD user exists. If exists, add additional characters to SAM
     
     Start-Countdown -Seconds 2 -Message "`n`nChecking if WindowsID $SAM already exists... "
        $CheckUser = Get-ADUser -LDAPFilter "(SAMAccountName=$SAM)"
            $Global:RequestedSAM = $SAM
    
	if ($NULL -eq $CheckUser){
        write-host "`nWindows ID: $SAM is not taken."  -ForegroundColor Green
		}else{

            while($NULL -ne $CheckUser) {
		        
                $EndChar++
                $Global:newSAM = $FirstName + $LastName.SubString($StartChar,$EndChar)  
            
            $CheckUser = Get-ADUser -LDAPFilter "(SAMAccountName=$newSAM)"
					
            }

            $Global:SAM = $newSAM 

            write-host "`nWindows ID: $RequestedSAM is taken. Username will be adjusted to $newSAM.`n" -ForegroundColor Red

    }
}	

Function Start-ValidateInformation{         

Start-Countdown -Seconds 2 -Message "Gathering table information..."

# Display information for validation

     $table = new-object psobject -Property @{
                   'EmployeeName' = "$Global:FirstName $LastName"
                   'Department' = $Global:DEPARTMENT
                   'Manager' = $Global:ManagerFull
                   'WebsiteID' = $Global:WEBACCESS
                   'MXIEID' = $Global:MXIEEXT
                   'Branch' = $Global:BRANCH
               }

	$table | Format-Table -AutoSize -Property EmployeeName,WebsiteID,MXIEID,Department,Manager

	# Verify information for accuracy

  Write-Host "Please verify that information is correct!" -ForegroundColor Red
     $IsCorrect = Read-Host "[Y] Yes, [N] No"

     switch($ISCorrect) {
        Y {
            Write-Host "`nInformation verified. Continuing script... " -ForegroundColor Green
			} 
        
		N {
					
             Get-UserInfo
             Start-IfExist
		     Start-ValidateInformation
			           
		}
	}
}

Function Start-CreateUser{

Start-Countdown -Seconds 2 -Message "Gathering table information..."
	     
    Write-Host "`n`n"

	$table = new-object psobject -Property @{
			'WindowsID' = $SAM
			'UserPrincipalName' = $UPN
			'DisplayName' = $DISPLAYNAME
			'Company' = $COMPANY
			'Container' = $OU
		    'Email' = $EmployeeEmail
               }
	$table | Format-Table -AutoSize -Property Email,WindowsID,UserPrincipalName
	$table | Format-Table -AutoSize -Property DisplayName,Company,Container

	Write-Host "`n`nReady to make user? " -NoNewLine -ForeGroundColor Yellow
			 $Selection = Read-Host "[Y] Create Account, [N] Abort"
	switch($Selection){		
		y{

      New-ADUser -Name $DISPLAYNAME -DisplayName $DISPLAYNAME -AccountPassword (ConvertTo-SecureString "$DefPWD" -AsPlainText -Force) -ChangePasswordAtLogon $FALSE -Company $COMPANY -Department $DEPARTMENT -Manager $Manager -Title $JOBTITLE -GivenName $FirstName -Surname $LastName -SamAccountName "$SAM" -EmailAddress "$EmployeeEmail" -Path $OU -UserPrincipalName $UPN -Enable $TRUE
		Start-Countdown -Seconds 2 -Message "Creating User Account..."
	
	# Set User Groups for account
    
	Write-Host "`nSelect Groups for $SAM " -ForegroundColor Yellow
    Start-Countdown -Seconds 2 -Message "Querying AD Groups... "
    $SGroups = Get-ADgroup -Filter * | Select-Object -Property Name | Out-GridView -PassThru | Select-Object -ExpandProperty Name 
         ForEach ($Group in $SGroups){ 
        Add-ADGroupMember -Members $SAM -Identity $Group
			} 

		}
		
		n{
            Write-Host "Aborting. Press any key to exit" -NoNewLine -ForeGroundColor Red
                $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
			CLS
			Start-UAW 
			}
		
		default {
            Write-Host "Invalid selection. Aborting. Press any key to exit" -NoNewLine -ForeGroundColor Red
                $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
			CLS
			Start-UAW                
		}
	}
}

Function Start-CreateSecondary {

	# Create Widgets account if desired

  Switch ($ISResponse) {
        Y {
            Write-Host "Widgets account has been created. `n" -foregroundcolor Green
            }
        
        N {
          Write-Host "`nContoso account created. Does this user also need an Widgets account? " -nonewline -ForeGroundColor Yellow
          $Global:SecondaryResponse = Read-Host "[Y] Yes, [N] No"
		}
	}
	    Switch ($SecondaryResponse) {
		    Y {
			    $Global:SecondaryEmail = "$FirstName.$LastName@Widgets.com"
			    $Global:SecondarySAM = "$SAM-IS"
			    $Global:SecondaryUPN = $SecondarySAM + “$DNSROOT”
                    
		    Start-Countdown -Seconds 2 -Message "Transferring information from Contoso account... "
			    Write-Host "`nSecondary account details collected. Select container for Widgets account... " -foregroundcolor Yellow
		
		    Start-Countdown -Seconds 2 -Message "Querying AD... "
		    $SecondaryOU = Get-ADOrganizationalUnit -Filter * | Select-Object -Property DistinguishedName | Out-GridView -PassThru | Select-Object -ExpandProperty DistinguishedName 
            
		    # Create Widgets user
                            
		    Start-Countdown -Seconds 2 -Message "Creating secondary Widgets account... "
			    New-ADUser -Name "$DISPLAYNAME (Widgets)" -DisplayName "$DISPLAYNAME (Widgets)" -AccountPassword (ConvertTo-SecureString "$DefPWD" -AsPlainText -Force) -ChangePasswordAtLogon $FALSE -Company "Widgets" -Department $DEPARTMENT -Manager $Manager -Title $JOBTITLE -GivenName $FirstName -Surname $LastName -SamAccountName "$SecondarySAM" -EmailAddress $SecondaryEmail -Path $SecondaryOU -UserPrincipalName "$SecondaryUPN" -Enable $TRUE
				
			
		    # Assign Widgets user groups
       
            Write-Host "`nSelect groups for $SecondarySAM.`n" -foregroundcolor Yellow                                 
		        Start-Countdown -Seconds 2 -Message "Querying AD Groups... "
			        $SGroups = Get-ADgroup -Filter * | Select-Object -Property Name | Out-GridView -PassThru | Select-Object -ExpandProperty Name 
			       	    ForEach ($Group in $SGroups){ 
			        	    Add-ADGroupMember -Members $SecondarySAM -Identity $Group
				    } 
			    }
		
		    N {
			    Write-Host "Skipping secondary account creation... `n" -foregroundcolor Red
			    }
	}
}

Function Start-CreateDrive{
		
	# Set up User Drive
	
	Write-Host "Would you like to set up User Drive? " -NoNewLine -ForeGroundColor Yellow 
			$Global:DriveResponse = Read-Host "[Y] Yes, [N] No"  

			switch ($DriveResponse) {
			 Y {
            # Set Drive location based on branch

                    switch($BRANCH){

                    1{
                    $Global:homedir="\\zeus\users\"
                    }

                    2{
                    $Global:homedir="\\10.0.2.225\users\"
                    }

                    3{
                    $Global:homedir="\\10.0.3.225\users\"
                    }

                    4{
                    $Global:homedir="\\10.0.4.225\users\"
                    }

                    8{
                    $Global:homedir="\\10.0.8.225\users\"
                    }

                    10{
                    $Global:homedir="\\10.0.10.225\users\"
                    }

                    11{
                    $Global:homedir="\\10.0.11.225\users\"
                    }

                    12{
                    $Global:homedir="\\10.0.12.225\users\"
                    }

                    16{
                    $Global:homedir="\\10.0.16.225\users\"
                    }
        }

			 
            Get-ADUser $SAM | % { Set-ADUser $_ -HomeDrive "U:" -HomeDirectory ($homedir + $SAM)
			    Write-Host "Creating drive.. `n" -NoNewLine -ForeGroundColor Yellow
				
				 Start-Countdown -Seconds 2 -Message "Creating & enabling User Drive... "
				  if (-not (Test-Path "$homedir$SAM")){ 
					
					  Start-Countdown -Seconds 2 -Message "Setting drive permissions... `n"
						$acl = Get-Acl (New-Item -Path $homedir -Name $SAM -ItemType Directory) 
 
			   # Make sure access rules inherited from parent folders. 
						$acl.SetAccessRuleProtection($false, $true) 
 
						$ace = "$domain\$SAM","FullControl", "ContainerInherit,ObjectInherit","None","Allow" 
						$objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($ace) 
						$acl.AddAccessRule($objACE) 
							 Set-ACL -Path "$homedir$SAM" -AclObject $acl 
								$SetDrive = $TRUE
								  } 
							   } 
									  
		}
			 N { 
				Write-Host "Skipping drive creation... `n" -NoNewLine -ForeGroundColor Red
				}
	}
}

Function Start-MakeMailbox{

# Set up Mailbox

    Write-Host "`nWould you like to enable user mailbox? " -NoNewLine -ForeGroundColor Yellow 
        $Global:MailResponse = Read-Host "[Y] Yes, [N] No"  

switch($MailResponse) {
    Y {
        switch($SecondaryResponse){
            Y {
            $UserCredential = Get-Credential -Credential "$env:USERDOMAIN\$env:USERNAME"
            $Session = New-PSSession -Credential $UserCredential -ConfigurationName Microsoft.Exchange -ConnectionUri "$ExchangeServerURL/powershell/" -Authentication Kerberos
                Import-PSSession $Session -DisableNameChecking | Out-Null
                      Enable-Mailbox -Identity $SAM -database "Contoso"
                      $SetMailbox = $TRUE 
                      Enable-Mailbox -Identity $SecondarySAM -database "Widgets"
                      $SetMailbox = $TRUE

                      # Sleep terminal to wait for Mailbox(s) to enable on server

                            Start-Countdown -Seconds 5 -Message "Enabling mailboxes... `n"
                            write-host "Mailboxes have been enabled. Continuing script... `n"
						}
			N {
				$UserCredential = Get-Credential -Credential "$env:USERDOMAIN\$env:USERNAME"
				$Session = New-PSSession -Credential $UserCredential -ConfigurationName Microsoft.Exchange -ConnectionUri "$ExchangeServerURL/PowerShell/" -Authentication Kerberos
				Import-PSSession $Session -DisableNameChecking | Out-Null
						Enable-Mailbox -Identity $SAM -database "Contoso"
						$SetMailbox = $TRUE  
                                              
            # Sleep terminal to wait for Mailbox(s) to enable on server

                Start-Countdown -Seconds 5 -Message "Enabling mailboxes... `n"
                write-host "Mailbox has been enabled. Continuing script... `n"
							}
						}
					}

			N {
			Write-Host "Skipping enabling mailbox... `n" -NoNewLine -ForeGroundColor Red 
						}
	}
}

Function Start-SendCredentials{
	
	# Send credentials via email
	Write-Host "`nWould you like to email user credentials to employee manager? " -NoNewLine -ForeGroundColor Yellow 
	  $Global:SendEmailResponse = Read-Host "[Y] Yes, [N] No"  

	switch ($SendEmailResponse){
		Y {
			switch($SecondaryResponse){
        
				Y {
            
				Write-Host "Please enter the Managers email address: " -NoNewLine -ForeGroundColor Yellow
					$ManagerEmail = Read-Host

	# Dual Account Email Body

			$DualEmail = @"
		<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta name=Generator content="Microsoft Word 15 (filtered medium)"><style><!--
		/* Font Definitions */
		@font-face
			{font-family:"Cambria Math";
			panose-1:2 4 5 3 5 4 6 3 2 4;}
		@font-face
			{font-family:Calibri;
			panose-1:2 15 5 2 2 2 4 3 2 4;}
		/* Style Definitions */
		p.MsoNormal, li.MsoNormal, div.MsoNormal
			{margin:0in;
			margin-bottom:.0001pt;
			font-size:11.0pt;
			font-family:"Calibri",sans-serif;}
		a:link, span.MsoHyperlink
			{mso-style-priority:99;
			color:#0563C1;
			text-decoration:underline;}
		a:visited, span.MsoHyperlinkFollowed
			{mso-style-priority:99;
			color:#954F72;
			text-decoration:underline;}
		p.msonormal0, li.msonormal0, div.msonormal0
			{mso-style-name:msonormal;
			mso-margin-top-alt:auto;
			margin-right:0in;
			mso-margin-bottom-alt:auto;
			margin-left:0in;
			font-size:11.0pt;
			font-family:"Calibri",sans-serif;}
		span.EmailStyle18
			{mso-style-type:personal-compose;
			font-family:"Calibri",sans-serif;
			color:windowtext;}
		.MsoChpDefault
			{mso-style-type:export-only;
			font-size:10.0pt;
			font-family:"Calibri",sans-serif;}
		@page WordSection1
			{size:8.5in 11.0in;
			margin:1.0in 1.0in 1.0in 1.0in;}
		div.WordSection1
			{page:WordSection1;}
		--></style><!--[if gte mso 9]><xml>
		<o:shapedefaults v:ext="edit" spidmax="1026" />
		</xml><![endif]--><!--[if gte mso 9]><xml>
		<o:shapelayout v:ext="edit">
		<o:idmap v:ext="edit" data="1" />
		</o:shapelayout></xml><![endif]--></head><body lang=EN-US link="#0563C1" vlink="#954F72"><div class=WordSection1><p class=MsoNormal>Hello $ManagerFirst,<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Accounts for $FirstName $LastName have been established and are as followed:<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Windows &amp; Outlook:<o:p></o:p></b></p><p class=MsoNormal><b><o:p>&nbsp;</o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b>Username: $SAM<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Email: $EmployeeEmail<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Default password: $DefPWD<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Secondary Widgets Account<o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b>Username: $SecondarySAM<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Email: $SecondaryEmail<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Default password: $DefPWD<o:p></o:p></p><p class=MsoNormal><b><o:p>&nbsp;</o:p></b></p><p class=MsoNormal><b><o:p>&nbsp;</o:p></b></p><p class=MsoNormal style='margin-left:1.0in'><i>Upon logging in the computer will prompt for a new password. This must be at least 8 characters long, include at least 1 number and at least 1 capital letter. Once logged in, opening Outlook 2016 will populate all accounts automatically. If you are prompted for a password enter the newly created password, not the default password. Windows &amp; Outlook will share the same account.<o:p></o:p></i></p><p class=MsoNormal><i><o:p>&nbsp;</o:p></i></p><p class=MsoNormal><i>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </i><b>MXIE Phone System:<o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b>Username: $MXIEEXT<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Default password: $MXIEDefPWD<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; URL: $MXIEURL<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Default voicemail pin: $MXIEDefPIN<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal style='margin-left:1.0in'><b>Voicemail message: On the bottom left corner of the Polycom phone screen you can click the button below VMAIL and follow the prompts to change your voicemail message.<o:p></o:p></b></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal style='margin-left:1.0in'><i>Your phone extension is the same as your username. This system will not automatically prompt to change the default pin or password. It is highly recommended that you change these settings by going to the top left corner and choosing File &gt; Change Password followed by File &gt; Change Pin. <o:p></o:p></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>RS&amp;I Website Account:<o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b>Username: $WEBACCESS<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>A temporary password will be provided to the newly created email account. Once you have entered the temporary password on </i><a href="https://www.contosoinc.com">https://www.contosoinc.com</a> <i>you will be prompted to change this password. This account functions on both the RS&amp;I website, as well as <a href="https://www.Widgets.com">https://www.Widgets.com</a></i><o:p></o:p></p></div></body></html>
"@
                                                
		# Send the email         
            
				Send-MailMessage -From $SendingEmail -To $ManagerEmail -Cc $SendingEmail,$EmployeeEmail -Subject "Credentials for $FirstName $LastName" -SmtpServer $SMTPServer -body ($DualEmail | Out-String) -bodyashtml
   
			}
                
			N {

					Write-Host "Please enter the email where credentials will be sent: " -NoNewLine -ForeGroundColor Yellow
						$ManagerEmail = Read-Host
        
	# Standard Email Body

			$StandardEmail = @"
		<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta name=Generator content="Microsoft Word 15 (filtered medium)"><style><!--
		/* Font Definitions */
		@font-face
			{font-family:"Cambria Math";
			panose-1:2 4 5 3 5 4 6 3 2 4;}
		@font-face
			{font-family:Calibri;
			panose-1:2 15 5 2 2 2 4 3 2 4;}
		/* Style Definitions */
		p.MsoNormal, li.MsoNormal, div.MsoNormal
			{margin:0in;
			margin-bottom:.0001pt;
			font-size:11.0pt;
			font-family:"Calibri",sans-serif;}
		a:link, span.MsoHyperlink
			{mso-style-priority:99;
			color:#0563C1;
			text-decoration:underline;}
		a:visited, span.MsoHyperlinkFollowed
			{mso-style-priority:99;
			color:#954F72;
			text-decoration:underline;}
		p.msonormal0, li.msonormal0, div.msonormal0
			{mso-style-name:msonormal;
			mso-margin-top-alt:auto;
			margin-right:0in;
			mso-margin-bottom-alt:auto;
			margin-left:0in;
			font-size:11.0pt;
			font-family:"Calibri",sans-serif;}
		span.EmailStyle18
			{mso-style-type:personal-compose;
			font-family:"Calibri",sans-serif;
			color:windowtext;}
		.MsoChpDefault
			{mso-style-type:export-only;
			font-size:10.0pt;
			font-family:"Calibri",sans-serif;}
		@page WordSection1
			{size:8.5in 11.0in;
			margin:1.0in 1.0in 1.0in 1.0in;}
		div.WordSection1
			{page:WordSection1;}
		--></style><!--[if gte mso 9]><xml>
		<o:shapedefaults v:ext="edit" spidmax="1026" />
		</xml><![endif]--><!--[if gte mso 9]><xml>
		<o:shapelayout v:ext="edit">
		<o:idmap v:ext="edit" data="1" />
		</o:shapelayout></xml><![endif]--></head><body lang=EN-US link="#0563C1" vlink="#954F72"><div class=WordSection1><p class=MsoNormal>Hello $ManagerFirst,<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Accounts for $FirstName $LastName have been established and are as followed:<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Windows &amp; Outlook:<o:p></o:p></b></p><p class=MsoNormal><b><o:p>&nbsp;</o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b>Username: $SAM<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Email: $EmployeeEmail<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Default password: $DefPWD<b><o:p></o:p></b></p><p class=MsoNormal><b><o:p>&nbsp;</o:p></b></p><p class=MsoNormal style='margin-left:1.0in'><i>Upon logging in the computer will prompt for a new password. This must be at least 8 characters long, include at least 1 number and at least 1 capital letter. Once logged in, opening Outlook 2016 will populate all accounts automatically. If you are prompted for a password enter the newly created password, not the default password. Windows &amp; Outlook will share the same account.<o:p></o:p></i></p><p class=MsoNormal><i><o:p>&nbsp;</o:p></i></p><p class=MsoNormal><i>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </i><b>MXIE Phone System:<o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b>Username: $MXIEEXT<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Default password: $MXIEDefPWD<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; URL: $MXIEURL<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Default voicemail pin: $MXIEDefPIN<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal style='margin-left:1.0in'><b>Voicemail message: On the bottom left corner of the Polycom phone screen you can click the button below VMAIL and follow the prompts to change your voicemail message.<o:p></o:p></b></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal style='margin-left:1.0in'><i>Your phone extension is the same as your username. This system will not automatically prompt to change the default pin or password. It is highly recommended that you change these settings by going to the top left corner and choosing File &gt; Change Password followed by File &gt; Change Pin. <o:p></o:p></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>RS&amp;I Website Account:<o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <o:p></o:p></b></p><p class=MsoNormal><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b>Username: $WEBACCESS<o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <o:p></o:p></p><p class=MsoNormal>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <i>A temporary password will be provided to the newly created email account. Once you have entered the temporary password on </i><a href="https://www.contosoinc.com">https://www.contosoinc.com</a> <i>you will be prompted to change this password.<o:p></o:p></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
"@

		 # Send the email
        
				  Send-MailMessage -From $SendingEmail -To $ManagerEmail -Cc $SendingEmail,$EmployeeEmail -Subject "Credentials for $FirstName $LastName" -SmtpServer $SMTPServer -body ($StandardEmail | out-string) -bodyashtml

					}
				}
			} 
			
		N {
			Write-Host "Skipping sending email... `n" -ForegroundColor Red
	   }
	}
}

Function Start-ChangePassword{

Start-Countdown -Seconds 2 -Message "Expiring temporary password... "

# Set user Password to change at next logon

switch($SecondaryResponse){

		Y{
			Get-ADUser -Identity "$SAM" | Set-ADUser -ChangePasswordAtLogon $TRUE
			Get-ADUser -Identity "$SecondarySAM" | Set-ADUser -ChangePasswordAtLogon $TRUE
		}
		
        N{
			Get-ADUser -Identity "$SAM" | Set-ADUser -ChangePasswordAtLogon $TRUE
		}
		
    }

}

#-------------------------------------------#

# Functions for interacting with AD Objects #

#-------------------------------------------#

Function Start-CheckAD{

		$checkADSPLASH = Get-Content -Path "$scriptDir\splashscreens\checkAD.splash" | Write-Host -ForegroundColor Cyan

Start-Countdown -Seconds 2 -Message "Querying AD... "  

# Set required variables


   $Date = [datetime]::Today.AddDays(-3)


# Assign AD objects to variables

   $ADComp = Get-ADComputer -Properties whencreated,name -Filter {whencreated -gt $Date} | Select-Object Name,whencreated | sort -Property Name 

   $ADUsers = Get-ADUser -Properties whencreated,SamAccountName -Filter {whencreated -gt $Date} | Select-Object SamAccountName,whencreated | sort -Property SamAccountName 


# Output information on screen

if ($NULL -eq $ADComp){

    Write-Host "`nNo new AD Computer Objects found" -ForegroundColor Green
    
    }else{

    Write-Host "`nNew AD Computer Objects have been found.`n" -ForegroundColor Red

    $AdComp | ft -Property Name,whencreated -AutoSize

    }

if ($NULL -eq $ADUsers){

    Write-Host "`nNo new AD User Objects found" -ForegroundColor Green
    
    }else{

    Write-Host "`nNew AD User Objects have been found.`n" -ForegroundColor Red

    $AdUsers | ft -Property SamAccountName,whencreated -AutoSize

    }

    Write-Host "`n`n                                      Press any key to return to the main screen" -ForegroundColor Yellow
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL

            CLS
            Start-ADME

}

Function Start-RemoveStaleADComputers{

	$stalecomputersSPLASH = Get-Content -Path "$scriptDir\splashscreens\staleCOMPUTERS.splash" | Write-Host -ForegroundColor Cyan

# Declare required variables

$Date = [datetime]::Today.AddDays(-365)

# Get ADObjects

$ADComputers = Get-ADComputer -Filter 'enabled -eq "true"' -Properties * | Select-Object Name,LastLogonDate,DistinguishedName | Where-Object {$_.LastLogonDate -lt $Date} | Sort-Object LastLogonDate

# Print table displaying object

$ADComputers | ft @{l="Object Name";e={$_.name}},@{l="Last Logon";e={$_.LastLogonDate}}

if($NULL -ne $ADComputers){

Write-Host "The following inactive AD computers have been found. Remove all? [Y] | [N]: " -NoNewline -ForegroundColor Yellow
    $Selection = Read-Host

     switch($Selection){

    Y{
        foreach ($Computers in $ADComputers){
        
        if(Test-Connection $Computers.name -Count 1 -ErrorAction SilentlyContinue){
          
          Write-Host "$($Computers.name) is live. Skipping to next object" -ForegroundColor Green
            
            }else{
                     
                      Write-Host "Removing $($Computers.name)" -ForegroundColor Red
                     
                      Remove-ADObject -Recucontosove -identity $Computers.DistinguishedName -Confirm:$false
           
                    }
                }
            }
   
    N{
        exit
    }
    
    Default{
                exit
            }
        }
    } elseif ($NULL -eq $ADComputers){
        Write-Host "No stale AD computers were found." -NoNewline -ForegroundColor Green
    }

 Write-Host "`n`n Script is finished. Stale AD computers have been purged. Press any key to return to the menu." -ForegroundColor Green
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
	CLS
	Start-ADME
}

Function Start-MoveStaleADUsers{

	$staleusersSPLASH = Get-Content -Path "$scriptDir\splashscreens\staleUSERS.splash" | Write-Host -ForegroundColor Cyan

# Declare required variables

$Date = [datetime]::Today.AddDays(-365)

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

$StaleOU = "OU=Stale Accounts,OU=contoso Miscellaneous User Accounts,DC=Olympus,DC=contosoinc,DC=com"

# Get ADObjects

$ADUsers = Get-ADUser -Filter 'enabled -eq "true"' -Properties * | Select-Object Name,SamAccountName,LastLogonDate,DistinguishedName | Where-Object {$_.LastLogonDate -lt $Date -and $_.Name -notlike 'IUSR_*' -and $_.Name -notlike 'IWAM_*' -and $_.Name -notlike 'HealthMailbox*' -and $_.Name -notlike 'VUSR_*'} | Sort-Object LastLogonDate

$ADUsers | Select-Object Name,SamAccountName,DistinguishedName | Out-File -FilePath "$ScriptDir\StaleUsers.txt"

# Print table displaying object

$ADUsers | ft @{l="Object Name";e={$_.name}},@{l="Last Logon";e={$_.LastLogonDate}}

if($NULL -ne $ADUsers){

Write-Host "The following inactive AD objects have been found. Move all? [Y] | [N]: " -NoNewline -ForegroundColor Yellow
    $Selection = Read-Host

     switch($Selection){

    Y{
        foreach ($Users in $ADUsers){
			
			Disable-ADAccount -Identity $Users.distinguishedname -Confirm:$false                           
			Move-ADObject -Identity $Users.distinguishedname -TargetPath $StaleOU
           
                    }
                }
   
    N{
        exit
    }
    
    Default{
                exit
            }
        }
    } elseif ($NULL -eq $ADUsers) {
        Write-Host "No Stale AD Objects were found." -NoNewline -ForegroundColor Green
    }

 Write-Host "`n`n Script is finished. Stale AD objects have been moved to Stale OU. Press any key to exit" -ForegroundColor Green
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL

}

Function Start-WACHandshake{

<#

This script will create a handshake for CredSSP between the Gateway and a Server you wish to manage with WAC. It will ask for the name of the Node 
and find it using the FQDN of that deivce per Get-ADComputer.

Written by Justin Sherley for use in Contoso | OmegaTechnologyServices@protonmail.com

#>

$wachsSPLASH = Get-Content -Path "$scriptDir\splashscreens\wachs.splash" | Write-Host -ForegroundColor Cyan

# Set required variables

$gateway = "Deployment" # Machine where Windows Admin Center is installed

# Get information from user

write-host "`nPlease enter the name of the machine you'd like to handshake with`n"
    $machine = read-host # Machine that you want to manage

# Complete handshake

$gatewayObject = Get-ADComputer -Identity $gateway
$machineObject = Get-ADComputer -Identity $machine

Set-ADComputer -Identity $machineObject -PrincipalsAllowedToDelegateToAccount $gatewayObject -Verbose

write-host "`nHandshake complete. Press enter to run command again. Enter Q to quit"
    $Response = Read-Host

    Switch($Response){

        Q{
            CLS
            Start-ACME
        }
        		
        default{
            CLS
            Start-WACHandshake
		}
    }
}

Function Start-CheckCluster{

	# Establish remote connection and sign in

	Write-Host "Please enter machine name: " -NoNewline -ForegroundColor Yellow
	$Machine = Read-Host
	$Creds = Get-Credential -Credential "$env:USERDOMAIN\$env:USERNAME"
 
	Function Start-CheckStorageJob{
		
	# Get Cluster Job status

    $ClusterJob = Get-StorageJob

    if($NULL -eq $ClusterJob){
        Write-Host "`nThere are currently no jobs running. Cluster is safe to reboot. Press any key to return to the menu" -ForegroundColor Green
           $Response = Read-Host
                
                switch($Response){
            
                default{
                    CLS
                    Start-ACME
                    }
                }
        
    } else {
        $ClusterJob | ft            
        Write-Host "Cluster Job is currently running. Press enter to check again. Enter Q to quit." -ForegroundColor Red
            $Response = Read-Host

        switch($Response){
            
        default{
            CLS
            Start-CheckCluster
        }

        Q{
            CLS
            Start-ACME
               }
           }
        }
    }

Invoke-Command -ComputerName $Machine -ScriptBlock ${Function:Start-CheckStorageJob} -credential $Creds

}

Function Start-ExtendClusterVolume{

<#

#>

	# Establish remote connection and sign in

	Write-Host "Please enter machine name: " -NoNewline -ForegroundColor Yellow
	$Machine = Read-Host
	$Creds = Get-Credential -Credential "$env:USERDOMAIN\$env:USERNAME"

Function Start-GetClusterInformation{

# Display total capacity of each pool available

Get-StorageTierSupportedSize -FriendlyName "Capacity" | FT @{l="FriendlyName";e={"Capacity"}} ,@{l="AvailableSpace(TB)";e={$_.TiecontosozeMax/1TB}}

Get-StorageTierSupportedSize -FriendlyName "Performance" | FT @{l="FriendlyName";e={"Performance"}} ,@{l="AvailableSpace(TB)";e={$_.TiecontosozeMax/1TB}}

# Display current Disk Pools

Get-VirtualDisk * | Sort FriendlyName | FT FriendlyName, OperationalStatus, HealthStatus, @{l='Size(GB)';E={$_.Size/1GB}},@{L='FootPrintOnPool(GB)';E={$_.FootprintOnPool/1GB}} -AutoSize


# Set variables

Write-Host "`nAvailable pools are above. Please enter the FriendlyName of the disk you'd like to resize: "
    $Global:vDisk = Read-Host

Write-Host "`nPlease enter total size you'd like this disk to be. E.X: 1024GB"
    $Global:Size = Read-Host


$Table = new-object psobject -Property @{
				'SelectedDisk' = $vDisk
				'DesiredSize' = $Size
}

$Table | Format-Table -AutoSize -Property SelectedDisk,DesiredSize

Write-Host "`n`nReady to execute? [Y] | [N]" -ForegroundColor Yellow
    $Global:Response = Read-Host

}


Function Start-ResizeDisk{

    switch($Response){

        Y{

        Get-VirtualDisk "$vDisk" | Get-StorageTier | Resize-StorageTier -Size $Size

        }

        N{

        Start-GetClusterInformation

        }

        Default{
            Write-Host "Invalid entry. Canceling... Press any key to return to menu"
                $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
            CLS
			Start-HCCM
        }

    }

}

Invoke-Command -ComputerName $Machine -ArgumentList $vDisk, $Size -ScriptBlock ${Function:Start-GetClusterInformation} -credential $Creds

Invoke-Command -ComputerName $Machine -ArgumentList $vDisk, $Size -ScriptBlock ${Function:Start-ResizeDisk} -credential $Creds

}

Function Start-DisableSMBv1{

<# 

This script will search for servers in a domain and then disable SMBv1 on each server.

Keep in mind that any machines running an OS prior to Windows 2008/Windows Vista require SMBv1 to work. Double check your environment
to ensure that there is no need for SMBv1.

Created by Justin Sherley for use in Contoso | omegatechnologyservices@protonmail.com

#>

<# 
	Coming: Display something less.. static on screen while we wait for Get-ADComputer...
#>

# Declare Functions

function Transpose-Data{
<#
	I'm unsure where this function originates, but it was recommended for use by on SpiceWorks by JitenSh.

	https://community.spiceworks.com/people/jitensh

	Forum link: https://community.spiceworks.com/posts/8512834

#>

    param(
        [String[]]$Names,
        [Object[][]]$Data
    )

    for($i = 0;; ++$i){
        $Props = [ordered]@{}
        for($j = 0; $j -lt $Data.Length; ++$j){
            if($i -lt $Data[$j].Length){
                $Props.Add($Names[$j], $Data[$j][$i])
            }
        }

        if(!$Props.get_Count()){
            break
        }

        [PSCustomObject]$Props
    }
}

function Start-Check2008Smb{

function Test-RegistryValue {

<#
	This function was taken from Johnathn Medd via his blog. I simply wrapped it with another function that makes it possible to invoke this remotely.

	Original Powershell Code located here: https://www.jonathanmedd.net/2014/02/testing-for-the-presence-of-a-registry-key-and-value.html

#>
    param (

     [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$Path,
     [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$Value
    )

    try {

    Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
        return $true
         
        } catch {

            return $false

    }
}

Test-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\' -Value 'SMB1'

}

# Gather Windows Servers from AD and put them into a variable

$Servers = Get-ADComputer -Filter {Operatingsystem -like '*windows server*'} -Properties OperatingSystem 

# Sort servers by Operating System

$Win2008 = $ser | where{$_.Operatingsystem -like '*2008*'} | select -ExpandProperty Name
$Win2012 = $ser | where{$_.Operatingsystem -like '*2012*'} | select -ExpandProperty Name
$Win2016 = $ser | where{$_.Operatingsystem -like '*2016*'} | select -ExpandProperty Name
$Win2019 = $ser | where{$_.Operatingsystem -like '*2019*'} | select -ExpandProperty Name

# Create table using Transpose-Data function

Transpose-Data Server2008,Server2012,Server2016,Server2019 $2008,$2012,$2016,$2019 | Format-Table -AutoSize | Out-File "$Global:ScriptDir\Files\Servers.txt"

	Get-Content -Path "$scriptDir\Files\Servers.txt" | Write-Host

# Prompt user and execute command based on the server

Write-Host "$Table `n`nThe above servers have been found. Are you ready to disable SMBv1 on each of these servers? [Y] | [N]: " -NoNewline -ForegroundColor Yellow
    $Selection = Read-Host

     switch($Selection){

        Y{
			$Cred = Get-Credential -Credential $env:USERDOMAIN\$env:USERNAME

				foreach ($computers in $Win2008){
			        
					if(Test-Connection $Computers.name -Count 1 -ErrorAction SilentlyContinue){
						
						Try{
						
						$RegKey = Invoke-Command -ComputerName $computers.name -Credential $Cred -ScriptBlock ${Function:Start-Check2008SMB} -ErrorAction Stop
				
							if($RegKey -eq $true){
								Write-Host "`nAttempting to disable SMBv1 on $($computers.name)...`n" -ForegroundColor Yellow
				
									Invoke-Command -ComputerName $computers.name -Credential $Cred -ScriptBlock { Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" SMB1 -Type DWORD -Value 0 –Force }

							} elseif($Regkey -ne $true) {
 
								Write-Host "SMBv1 is already disabled on $($computers.name)" -ForegroundColor Green
                
							}
				
							} catch [System.Management.Automation.RuntimeException] {

								Write-Host "Could not remotely connect to $($computers.name). Check WinRM & Firewall settings on target machine."
				
							} catch {

								Write-Host "Uknown error. Error caught is: $($_.Exception.GetType().FullName)"
							}
			} else {

				 Write-Host "$($Computers.name) is not alive. Skipping to next object" -ForegroundColor Green
			
			}
		}

				foreach ($computers in $Win2012 -and $Win2016 -and $Win2019){

					if(Test-Connection $Computers.name -Count 1 -ErrorAction SilentlyContinue){
                    
						Try{
            
							$SMB2012 = Invoke-Command -ComputerName $computers.name -Credential $Cred -ScriptBlock { Get-WindowsFeature FS-SMB1 } -ErrorAction Stop
			
								if($SMB2012.Installed -eq $true){
					
								Write-Host "`nAttempting to disable SMBv1 on $($computers.name)...`n" -ForegroundColor Yellow
				
								Invoke-Command -ComputerName $computers.name -Credential $Cred -ScriptBlock { Remove-WindowsFeature FS-SMB1 }

								} elseif($Feature.Installed -ne $true) {
 
								Write-Host "SMBv1 is already disabled on $($computers.name)" -ForegroundColor Green
                
								}
				
								} catch [System.Management.Automation.RuntimeException] {

								Write-Host "Could not remotely connect to $($computers.name). Check WinRM & Firewall settings on target machine."
				
								} catch {

								Write-Host "Uknown error. Error caught is: $($_.Exception.GetType().FullName)"
								}
								} else {

								Write-Host "$($Computers.name) is not alive. Skipping to next object" -ForegroundColor Green
			}
		}
	}
     
	N{

        Write-Host "Script aborted.. Press any key to exit the shell."
            $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL

        }

        Default{
          Write-Host "Invalid entry. Script aborted.. Press any key to exit the shell."
            $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
        }
    }

Write-Host "`n`nScript has finished.. Press any key to exit the shell." -ForegroundColor Green
	$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL

}

