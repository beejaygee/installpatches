Function Get-WUInstall
{

	#>
	[OutputType('PSWindowsUpdate.WUInstall')]
	[CmdletBinding(
		SupportsShouldProcess=$True,
		ConfirmImpact="High"
	)]	
	Param
	(
		#Pre search criteria
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[ValidateSet("Driver", "Software")]
		[String]$UpdateType="",
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String[]]$UpdateID,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[Int]$RevisionNumber,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String[]]$CategoryIDs,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[Switch]$IsInstalled,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[Switch]$IsHidden,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[Switch]$WithHidden,
		[String]$Criteria,
		[Switch]$ShowSearchCriteria,
		
		#Post search criteria
        [ValidateSet('Critical Updates', 'Definition Updates', 'Drivers', 'Feature Packs', 'Security Updates', 'Service Packs', 'Tools', 'Update Rollups', 'Updates', 'Upgrades', 'Microsoft')]
        [String[]]$RootCategories,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String[]]$Category="",
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String[]]$KBArticleID,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String]$Title,
		[parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet("Critical", "Important", "Moderate", "Low", "Unspecified", "")]
		[String[]]$Severity,
		
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String[]]$NotCategory="",
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String[]]$NotKBArticleID,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[String]$NotTitle,
		[parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet("Critical", "Important", "Moderate", "Low", "Unspecified", "")]
		[String[]]$NotSeverity,
        [Int]$MaxSize,
        [Int]$MinSize,
        		
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[Alias("Silent")]
		[Switch]$IgnoreUserInput,
		[parameter(ValueFromPipelineByPropertyName=$true)]
		[Switch]$IgnoreRebootRequired,
		
		#Connection options
		[String]$ServiceID,
		[Switch]$WindowsUpdate,
		[Switch]$MicrosoftUpdate,
		
		#Mode options
		[Switch]$ListOnly,
		[Switch]$DownloadOnly,
		[Alias("All")]
		[Switch]$AcceptAll,
		[Switch]$AutoReboot,
		[Switch]$IgnoreReboot,
		[Switch]$AutoSelectOnly,
		[Switch]$Debuger
	)

	Begin
	{
		If($PSBoundParameters['Debuger'])
		{
			$DebugPreference = "Continue"
		} #End If $PSBoundParameters['Debuger']
		
		$User = [Security.Principal.WindowsIdentity]::GetCurrent()
		$Role = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

		if(!$Role)
		{
			Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."	
		} #End If !$Role
	}

	Process
	{
		#region	STAGE 0	
		######################################
		# Start STAGE 0: Prepare environment #
		######################################
		
		Write-Debug "STAGE 0: Prepare environment"
		If($IsInstalled)
		{
			$ListOnly = $true
			Write-Debug "Change to ListOnly mode"
		} #End If $IsInstalled

		Write-Debug "Check reboot status only for local instance"
		Try
		{
			$objSystemInfo = New-Object -ComObject "Microsoft.Update.SystemInfo"	
			If($objSystemInfo.RebootRequired)
			{
				Write-Warning "Reboot is required to continue"
				If($AutoReboot)
				{
					Restart-Computer -Force
				} #End If $AutoReboot

				If(!$ListOnly)
				{
					Return
				} #End If !$ListOnly	
				
			} #End If $objSystemInfo.RebootRequired
		} #End Try
		Catch
		{
			Write-Warning "Support local instance only, Continue..."
		} #End Catch
		
		Write-Debug "Set number of stage"
		If($ListOnly)
		{
			$NumberOfStage = 2
		} #End $ListOnly
		ElseIf($DownloadOnly)
		{
			$NumberOfStage = 3
		} #End Else $ListOnly If $DownloadOnly
		Else
		{
			$NumberOfStage = 4
		} #End Else $DownloadOnly
		
		####################################			
		# End STAGE 0: Prepare environment #
		####################################
		#endregion
		
		#region	STAGE 1
		###################################
		# Start STAGE 1: Get updates list #
		###################################			
		
		Write-Debug "STAGE 1: Get updates list"
		Write-Debug "Create Microsoft.Update.ServiceManager object"
		$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" 
		
		Write-Debug "Create Microsoft.Update.Session object"
		$objSession = New-Object -ComObject "Microsoft.Update.Session" 
		
		Write-Debug "Create Microsoft.Update.Session.Searcher object"
		$objSearcher = $objSession.CreateUpdateSearcher()

		If($WindowsUpdate)
		{
			Write-Debug "Set source of updates to Windows Update"
			$objSearcher.ServerSelection = 2
			$serviceName = "Windows Update"
		} #End If $WindowsUpdate
		ElseIf($MicrosoftUpdate)
		{
			Write-Debug "Set source of updates to Microsoft Update"
			$serviceName = $null
            if($MicrosoftUpdate)
            {
                if((Get-WUServiceManager -WarningAction SilentlyContinue).Name -notcontains "Microsoft Update") 
                { 
                    Write-Verbose "Can't find registered service Microsoft Update. Trying to add..."
                    Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d -Confirm:$False -ErrorAction SilentlyContinue
                }#End (Get-WUServiceManager -WarningAction SilentlyContinue).Name -notcontains "Microsoft Update"
            }#End $MicrosoftUpdate

			Foreach ($objService in $objServiceManager.Services) 
			{
				If($objService.Name -eq "Microsoft Update")
				{
					$objSearcher.ServerSelection = 3
					$objSearcher.ServiceID = $objService.ServiceID
					$serviceName = $objService.Name
					Break
				}#End If $objService.Name -eq "Microsoft Update"
			}#End ForEach $objService in $objServiceManager.Services
			
			If(-not $serviceName)
			{
				Write-Warning "Can't find registered service Microsoft Update. Use Get-WUServiceManager to get registered service."
				Return
			}#Enf If -not $serviceName
		} #End Else $WindowsUpdate If $MicrosoftUpdate
		Else
		{
			Foreach ($objService in $objServiceManager.Services) 
			{
				If($ServiceID)
				{
					If($objService.ServiceID -eq $ServiceID)
					{
						$objSearcher.ServiceID = $ServiceID
						$objSearcher.ServerSelection = 3
						$serviceName = $objService.Name
						Break
					} #End If $objService.ServiceID -eq $ServiceID
				} #End If $ServiceID
				Else
				{
					If($objService.IsDefaultAUService -eq $True)
					{
						$serviceName = $objService.Name
						Break
					} #End If $objService.IsDefaultAUService -eq $True
				} #End Else $ServiceID
			} #End Foreach $objService in $objServiceManager.Services
		} #End Else $MicrosoftUpdate
		Write-Debug "Set source of updates to $serviceName"
		
		Write-Verbose "Connecting to $serviceName server. Please wait..."
		Try
		{
			$search = ""
			
			If($Criteria)
			{
				$search = $Criteria
			} #End If $Criteria
			Else
			{
				If($IsInstalled) 
				{
					$search = "IsInstalled = 1"
					Write-Debug "Set pre search criteria: IsInstalled = 1"
				} #End If $IsInstalled
				Else
				{
					$search = "IsInstalled = 0"	
					Write-Debug "Set pre search criteria: IsInstalled = 0"
				} #End Else $IsInstalled
				
				If($UpdateType -ne "")
				{
					Write-Debug "Set pre search criteria: Type = $UpdateType"
					$search += " and Type = '$UpdateType'"
				} #End If $UpdateType -ne ""					
				
				If($UpdateID)
				{
					Write-Debug "Set pre search criteria: UpdateID = '$([string]::join(", ", $UpdateID))'"
					$tmp = $search
					$search = ""
					$LoopCount = 0
					Foreach($ID in $UpdateID)
					{
						If($LoopCount -gt 0)
						{
							$search += " or "
						} #End If $LoopCount -gt 0
						If($RevisionNumber)
						{
							Write-Debug "Set pre search criteria: RevisionNumber = '$RevisionNumber'"	
							$search += "($tmp and UpdateID = '$ID' and RevisionNumber = $RevisionNumber)"
						} #End If $RevisionNumber
						Else
						{
							$search += "($tmp and UpdateID = '$ID')"
						} #End Else $RevisionNumber
						$LoopCount++
					} #End Foreach $ID in $UpdateID
				} #End If $UpdateID

				If($CategoryIDs)
				{
					Write-Debug "Set pre search criteria: CategoryIDs = '$([string]::join(", ", $CategoryIDs))'"
					$tmp = $search
					$search = ""
					$LoopCount =0
					Foreach($ID in $CategoryIDs)
					{
						If($LoopCount -gt 0)
						{
							$search += " or "
						} #End If $LoopCount -gt 0
						$search += "($tmp and CategoryIDs contains '$ID')"
						$LoopCount++
					} #End Foreach $ID in $CategoryIDs
				} #End If $CategoryIDs
				
				If($IsHidden) 
				{
					Write-Debug "Set pre search criteria: IsHidden = 1"
					$search += " and IsHidden = 1"	
				} #End If $IsNotHidden
				ElseIf($WithHidden) 
				{
					Write-Debug "Set pre search criteria: IsHidden = 1 and IsHidden = 0"
				} #End ElseIf $WithHidden
				Else
				{
					Write-Debug "Set pre search criteria: IsHidden = 0"
					$search += " and IsHidden = 0"	
				} #End Else $WithHidden
				
				#Don't know why every update have RebootRequired=false which is not always true
				If($IgnoreRebootRequired) 
				{
					Write-Debug "Set pre search criteria: RebootRequired = 0"
					$search += " and RebootRequired = 0"	
				} #End If $IgnoreRebootRequired
			} #End Else $Criteria
			
			Write-Debug "Search criteria is: $search"
			
			If($ShowSearchCriteria)
			{
				Write-Output $search
			} #End If $ShowSearchCriteria
			
			$objResults = $objSearcher.Search($search)
		} #End Try
		Catch
		{
			If($_ -match "HRESULT: 0x80072EE2")
			{
				Write-Warning "Probably you don't have connection to Windows Update server"
			} #End If $_ -match "HRESULT: 0x80072EE2"
			Return
		} #End Catch

		$objCollectionUpdate = New-Object -ComObject "Microsoft.Update.UpdateColl" 
		
		$NumberOfUpdate = 1
		$UpdateCollection = @()
		$UpdatesExtraDataCollection = @{}
		$PreFoundUpdatesToDownload = $objResults.Updates.count
		Write-Verbose "Found [$PreFoundUpdatesToDownload] Updates in pre search criteria"				

        if($RootCategories)
        {
            $RootCategoriesCollection = @()
            foreach($RootCategory in $RootCategories)
            {
                switch ($RootCategory) 
                { 
                    "Critical Updates" {$CatID = 0} 
                    "Definition Updates"{$CatID = 1} 
                    "Drivers"{$CatID = 2} 
                    "Feature Packs"{$CatID = 3} 
                    "Security Updates"{$CatID = 4} 
                    "Service Packs"{$CatID = 5} 
                    "Tools"{$CatID = 6} 
                    "Update Rollups"{$CatID = 7} 
                    "Updates"{$CatID = 8} 
                    "Upgrades"{$CatID = 9} 
                    "Microsoft"{$CatID = 10} 
                } #End switch $RootCategory
                Try { $RootCategoriesCollection += $objResults.RootCategories.item($CatID).Updates } Catch { Write-Error "RootCategiries Updates are empty. Use classic filters." -ErrorAction Stop }
            } #End foreach $RootCategory in $RootCategories
            $objResults = New-Object -TypeName psobject -Property @{Updates = $RootCategoriesCollection}
        } #End if $RootCategories

		Foreach($Update in $objResults.Updates)
		{	
			$UpdateAccess = $true

			Write-Debug "Convert size"
			Switch($Update.MaxDownloadSize)
			{
				{[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
				{[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }  
				{[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }    
				{[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
				default { $size = $_+"B" }
			} #End Switch

			Write-Progress -Activity "Post search updates for $Computer" -Status "[$NumberOfUpdate/$PreFoundUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$PreFoundUpdatesToDownload * 100))
			Write-Debug "Set post search criteria: $($Update.Title)"
			
			If($Category -ne "")
			{
				$UpdateCategories = $Update.Categories | Select-Object Name
				Write-Debug "Set post search criteria: Categories = '$([string]::join(", ", $Category))'"	
				Foreach($Cat in $Category)
				{
					If(!($UpdateCategories -match $Cat))
					{
						Write-Debug "UpdateAccess: false"
						$UpdateAccess = $false
					} #End If !($UpdateCategories -match $Cat)
					Else
					{
						$UpdateAccess = $true
						Break
					} #End Else !($UpdateCategories -match $Cat)
				} #End Foreach $Cat in $Category	
			} #End If $Category -ne ""

			If($NotCategory -ne "" -and $UpdateAccess -eq $true)
			{
				$UpdateCategories = $Update.Categories | Select-Object Name
				Write-Debug "Set post search criteria: NotCategories = '$([string]::join(", ", $NotCategory))'"	
				Foreach($Cat in $NotCategory)
				{
					If($UpdateCategories -match $Cat)
					{
						Write-Debug "UpdateAccess: false"
						$UpdateAccess = $false
						Break
					} #End If $UpdateCategories -match $Cat
				} #End Foreach $Cat in $NotCategory	
			} #End If $NotCategory -ne "" -and $UpdateAccess -eq $true					
			
			If($KBArticleID -ne $null -and $UpdateAccess -eq $true)
			{
				Write-Debug "Set post search criteria: KBArticleIDs = '$([string]::join(", ", $KBArticleID))'"
				If(!($KBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs))
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If !($KBArticleID -match $Update.KBArticleIDs)								
			} #End If $KBArticleID -ne $null -and $UpdateAccess -eq $true

			If($NotKBArticleID -ne $null -and $UpdateAccess -eq $true)
			{
				Write-Debug "Set post search criteria: NotKBArticleIDs = '$([string]::join(", ", $NotKBArticleID))'"
				If($NotKBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs)
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If$NotKBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs					
			} #End If $NotKBArticleID -ne $null -and $UpdateAccess -eq $true
			
			If($Title -and $UpdateAccess -eq $true)
			{
				Write-Debug "Set post search criteria: Title = '$Title'"
				If($Update.Title -notmatch $Title)
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If $Update.Title -notmatch $Title
			} #End If $Title -and $UpdateAccess -eq $true

			If($NotTitle -and $UpdateAccess -eq $true)
			{
				Write-Debug "Set post search criteria: NotTitle = '$NotTitle'"
				If($Update.Title -match $NotTitle)
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If $Update.Title -notmatch $NotTitle
			} #End If $NotTitle -and $UpdateAccess -eq $true

			If($Severity -and $UpdateAccess -eq $true)
			{
				if($Severity -contains "Unspecified") { $Severity += "" } 
                Write-Debug "Set post search criteria: Severity = '$Severity'"
				If($Severity -notcontains [String]$Update.MsrcSeverity)
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If $Severity -notcontains $Update.MsrcSeverity
			} #End If $Severity -and $UpdateAccess -eq $true

			If($NotSeverity -and $UpdateAccess -eq $true)
			{
				if($NotSeverity -contains "Unspecified") { $NotSeverity += "" } 
                Write-Debug "Set post search criteria: NotSeverity = '$NotSeverity'"
				If($NotSeverity -contains [String]$Update.MsrcSeverity)
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If $NotSeverity -contains $Update.MsrcSeverity
			} #End If $NotSeverity -and $UpdateAccess -eq $true

			If($MaxSize -and $UpdateAccess -eq $true)
			{
                Write-Debug "Set post search criteria: MaxDownloadSize <= '$MaxSize'"
				If($MaxSize -le $Update.MaxDownloadSize)
				{
				    Write-Debug "UpdateAccess: false"
				    $UpdateAccess = $false
			    } #End If $MaxSize -le $Update.MaxDownloadSize
		    } #End If $MaxSize -and $UpdateAccess -eq $true

			If($MinSize -and $UpdateAccess -eq $true)
			{
                Write-Debug "Set post search criteria: MaxDownloadSize >= '$MinSize'"
			    If($MinSize -ge $Update.MaxDownloadSize)
			    {
			        Write-Debug "UpdateAccess: false"
			        $UpdateAccess = $false
			    } #End If $MinSize -ge $Update.MaxDownloadSize
			} #End If $MinSize -and $UpdateAccess -eq $true
			
			If($IgnoreUserInput -and $UpdateAccess -eq $true)
			{
				Write-Debug "Set post search criteria: CanRequestUserInput"
				If($Update.InstallationBehavior.CanRequestUserInput -eq $true)
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If $Update.InstallationBehavior.CanRequestUserInput -eq $true
			} #End If $IgnoreUserInput -and $UpdateAccess -eq $true

			If($IgnoreRebootRequired -and $UpdateAccess -eq $true) 
			{
				Write-Debug "Set post search criteria: RebootBehavior"
				If($Update.InstallationBehavior.RebootBehavior -ne 0)
				{
					Write-Debug "UpdateAccess: false"
					$UpdateAccess = $false
				} #End If $Update.InstallationBehavior.RebootBehavior -ne 0	
			} #End If $IgnoreRebootRequired -and $UpdateAccess -eq $true

			If($UpdateAccess -eq $true)
			{
				Write-Debug "Convert KBArticleIDs"
				If($Update.KBArticleIDs -ne "")    
				{
					$KB = "KB"+$Update.KBArticleIDs
				} #End If $Update.KBArticleIDs -ne ""
				Else 
				{
					$KB = ""
				} #End Else $Update.KBArticleIDs -ne ""
				
				If($ListOnly)
				{
					$Status = ""
					If($Update.IsDownloaded)    {$Status += "D"} else {$status += "-"}
					If($Update.IsInstalled)     {$Status += "I"} else {$status += "-"}
					If($Update.IsMandatory)     {$Status += "M"} else {$status += "-"}
					If($Update.IsHidden)        {$Status += "H"} else {$status += "-"}
					If($Update.IsUninstallable) {$Status += "U"} else {$status += "-"}
					If($Update.IsBeta)          {$Status += "B"} else {$status += "-"} 
	
					Add-Member -InputObject $Update -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME
					Add-Member -InputObject $Update -MemberType NoteProperty -Name KB -Value $KB
					Add-Member -InputObject $Update -MemberType NoteProperty -Name Size -Value $size
					Add-Member -InputObject $Update -MemberType NoteProperty -Name Status -Value $Status
					Add-Member -InputObject $Update -MemberType NoteProperty -Name X -Value 1
					
					$Update.PSTypeNames.Clear()
					$Update.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
					$UpdateCollection += $Update
				} #End If $ListOnly
				Else
				{
					$objCollectionUpdate.Add($Update) | Out-Null
					$UpdatesExtraDataCollection.Add($Update.Identity.UpdateID,@{KB = $KB; Size = $size})
				} #End Else $ListOnly
			} #End If $UpdateAccess -eq $true
			
			$NumberOfUpdate++
		} #End Foreach $Update in $objResults.Updates				
		Write-Progress -Activity "[1/$NumberOfStage] Post search updates" -Status "Completed" -Completed
		
		If($ListOnly)
		{
			$FoundUpdatesToDownload = $UpdateCollection.count
		} #End If $ListOnly
		Else
		{
			$FoundUpdatesToDownload = $objCollectionUpdate.count				
		} #End Else $ListOnly
		Write-Verbose "Found [$FoundUpdatesToDownload] Updates in post search criteria"
		
		If($FoundUpdatesToDownload -eq 0)
		{
			Return
		} #End If $FoundUpdatesToDownload -eq 0
		
		If($ListOnly)
		{
			Write-Debug "Return only list of updates"
			Return $UpdateCollection				
		} #End If $ListOnly

		#################################
		# End STAGE 1: Get updates list #
		#################################
		#endregion
		

		If(!$ListOnly) 
		{
			#region	STAGE 2
			#################################
			# Start STAGE 2: Choose updates #
			#################################
			
			Write-Debug "STAGE 2: Choose updates"			
			$NumberOfUpdate = 1
			$logCollection = @()
			
			$objCollectionChoose = New-Object -ComObject "Microsoft.Update.UpdateColl"

			Foreach($Update in $objCollectionUpdate)
			{	
				$size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
				Write-Progress -Activity "[2/$NumberOfStage] Choose updates" -Status "[$NumberOfUpdate/$FoundUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$FoundUpdatesToDownload * 100))
				Write-Debug "Show update to accept: $($Update.Title)"
				
				If($AcceptAll)
				{
					$Status = "Accepted"

					If($Update.EulaAccepted -eq 0)
					{ 
						Write-Debug "Accept Eula"
						$Update.AcceptEula() 
					} #End If $Update.EulaAccepted -eq 0
			
					Write-Debug "Add update to collection"
					$objCollectionChoose.Add($Update) | Out-Null
				} #End If $AcceptAll
				ElseIf($AutoSelectOnly)  
				{  
					If($Update.AutoSelectOnWebsites)  
					{  
						$Status = "Accepted"  
						If($Update.EulaAccepted -eq 0)  
						{  
							Write-Debug "Accept Eula"  
							$Update.AcceptEula()  
						} #End If $Update.EulaAccepted -eq 0  
  
						Write-Debug "Add update to collection"  
						$objCollectionChoose.Add($Update) | Out-Null  
					} #End If $Update.AutoSelectOnWebsites 
					Else  
					{  
						$Status = "Rejected"  
					} #End Else $Update.AutoSelectOnWebsites
				} #End ElseIf $AutoSelectOnly
				Else
				{
					If($pscmdlet.ShouldProcess($Env:COMPUTERNAME,"$($Update.Title)[$size]?")) 
					{
						$Status = "Accepted"
						
						If($Update.EulaAccepted -eq 0)
						{ 
							Write-Debug "Accept Eula"
							$Update.AcceptEula() 
						} #End If $Update.EulaAccepted -eq 0
				
						Write-Debug "Add update to collection"
						$objCollectionChoose.Add($Update) | Out-Null 
					} #End If $pscmdlet.ShouldProcess($Env:COMPUTERNAME,"$($Update.Title)[$size]?")
					Else
					{
						$Status = "Rejected"
					} #End Else $pscmdlet.ShouldProcess($Env:COMPUTERNAME,"$($Update.Title)[$size]?")
				} #End Else $AutoSelectOnly
				
				Write-Debug "Add to log collection"
				$log = New-Object PSObject -Property @{
					Title = $Update.Title
					KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
					Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
					Status = $Status
					X = 2
				} #End PSObject Property
				
				$log.PSTypeNames.Clear()
				$log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
				
				$logCollection += $log
				
				$NumberOfUpdate++
			} #End Foreach $Update in $objCollectionUpdate
			Write-Progress -Activity "[2/$NumberOfStage] Choose updates" -Status "Completed" -Completed
			
			#Write-Debug "Show log collection"
			#$logCollection
			
			$AcceptUpdatesToDownload = $objCollectionChoose.count
			Write-Verbose "Accept [$AcceptUpdatesToDownload] Updates to Download"
			
			If($AcceptUpdatesToDownload -eq 0)
			{
				Return
			} #End If $AcceptUpdatesToDownload -eq 0	
				
			###############################
			# End STAGE 2: Choose updates #
			###############################
			#endregion
			
			#region STAGE 3
			###################################
			# Start STAGE 3: Download updates #
			###################################
			
			Write-Debug "STAGE 3: Download updates"
			$NumberOfUpdate = 1
			$objCollectionDownload = New-Object -ComObject "Microsoft.Update.UpdateColl" 

			Foreach($Update in $objCollectionChoose)
			{
                $size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size				
                Write-Progress -Activity "[3/$NumberOfStage] Downloading updates" -Status "[$NumberOfUpdate/$AcceptUpdatesToDownload] $($Update.Title) $size" -PercentComplete ([int]($NumberOfUpdate/$AcceptUpdatesToDownload * 100))
				Write-Debug "Show update to download: $($Update.Title)"
				
				Write-Debug "Send update to download collection"
				$objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
				$objCollectionTmp.Add($Update) | Out-Null
					
				$Downloader = $objSession.CreateUpdateDownloader() 
				$Downloader.Updates = $objCollectionTmp
				Try
				{
					Write-Debug "Try download update"
					$DownloadResult = $Downloader.Download()
				} #End Try
				Catch
				{
					If($_ -match "HRESULT: 0x80240044")
					{
						Write-Warning "Your security policy don't allow a non-administator identity to perform this task"
					} #End If $_ -match "HRESULT: 0x80240044"
					
					Return
				} #End Catch 
				
				Write-Debug "Check ResultCode"
				Switch -exact ($DownloadResult.ResultCode)
				{
					0   { $Status = "NotStarted" }
					1   { $Status = "InProgress" }
					2   { $Status = "Downloaded" }
					3   { $Status = "DownloadedWithErrors" }
					4   { $Status = "Failed" }
					5   { $Status = "Aborted" }
				} #End Switch
				
				Write-Debug "Add to log collection"
				$log = New-Object PSObject -Property @{
					Title = $Update.Title
					KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
					Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
					Status = $Status
					X = 3
				} #End PSObject Property
				
				$log.PSTypeNames.Clear()
				$log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
				
				$log
				
				If($DownloadResult.ResultCode -eq 2)
				{
					Write-Debug "Downloaded then send update to next stage"
					$objCollectionDownload.Add($Update) | Out-Null
				} #End If $DownloadResult.ResultCode -eq 2
				
				$NumberOfUpdate++
				
			} #End Foreach $Update in $objCollectionChoose
			Write-Progress -Activity "[3/$NumberOfStage] Downloading updates" -Status "Completed" -Completed

			$ReadyUpdatesToInstall = $objCollectionDownload.count
			Write-Verbose "Downloaded [$ReadyUpdatesToInstall] Updates to Install"
		
			If($ReadyUpdatesToInstall -eq 0)
			{
				Return
			} #End If $ReadyUpdatesToInstall -eq 0
		

			#################################
			# End STAGE 3: Download updates #
			#################################
			#endregion
			
			If(!$DownloadOnly)
			{
				#region	STAGE 4
				##################################
				# Start STAGE 4: Install updates #
				##################################
				
				Write-Debug "STAGE 4: Install updates"
				$NeedsReboot = $false
				$NumberOfUpdate = 1
				
				#install updates	
				Foreach($Update in $objCollectionDownload)
				{   
					Write-Progress -Activity "[4/$NumberOfStage] Installing updates" -Status "[$NumberOfUpdate/$ReadyUpdatesToInstall] $($Update.Title)" -PercentComplete ([int]($NumberOfUpdate/$ReadyUpdatesToInstall * 100))
					Write-Debug "Show update to install: $($Update.Title)"
					
					Write-Debug "Send update to install collection"
					$objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
					$objCollectionTmp.Add($Update) | Out-Null
					
					$objInstaller = $objSession.CreateUpdateInstaller()
					$objInstaller.Updates = $objCollectionTmp
						
					Try
					{
						Write-Debug "Try install update"
						$InstallResult = $objInstaller.Install()
					} #End Try
					Catch
					{
						If($_ -match "HRESULT: 0x80240044")
						{
							Write-Warning "Your security policy don't allow a non-administator identity to perform this task"
						} #End If $_ -match "HRESULT: 0x80240044"
						
						Return
					} #End Catch
					
					If(!$NeedsReboot) 
					{ 
						Write-Debug "Set instalation status RebootRequired"
						$NeedsReboot = $installResult.RebootRequired 
					} #End If !$NeedsReboot
					
					Switch -exact ($InstallResult.ResultCode)
					{
						0   { $Status = "NotStarted"}
						1   { $Status = "InProgress"}
						2   { $Status = "Installed"}
						3   { $Status = "InstalledWithErrors"}
						4   { $Status = "Failed"}
						5   { $Status = "Aborted"}
					} #End Switch
				   
					Write-Debug "Add to log collection"
					$log = New-Object PSObject -Property @{
						Title = $Update.Title
						KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
						Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
						Status = $Status
						X = 4
					} #End PSObject Property
					
					$log.PSTypeNames.Clear()
					$log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
					
					$log
				
					$NumberOfUpdate++
				} #End Foreach $Update in $objCollectionDownload
				Write-Progress -Activity "[4/$NumberOfStage] Installing updates" -Status "Completed" -Completed
				
				If($NeedsReboot)
				{
					If($AutoReboot)
					{
						Restart-Computer -Force
					} #End If $AutoReboot
					ElseIf($IgnoreReboot)
					{
						Return "Reboot is required, but do it manually."
					} #End Else $AutoReboot If $IgnoreReboot
					Else
					{
						$Reboot = Read-Host "Reboot is required. Do it now ? [Y/N]"
						If($Reboot -eq "Y")
						{
							Restart-Computer -Force
						} #End If $Reboot -eq "Y"
						
					} #End Else $IgnoreReboot	
					
				} #End If $NeedsReboot

				################################
				# End STAGE 4: Install updates #
				################################
				#endregion
			} #End If !$DownloadOnly
		} #End !$ListOnly
	} #End Process
	
	End{}		
}
