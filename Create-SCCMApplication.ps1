[CmdletBinding(SupportsShouldProcess = $true)]
param (
	[parameter(Mandatory = $true, HelpMessage = "Name of application")]
	[ValidateNotNullOrEmpty()]
	[string]$ApplicationName,

    [parameter(Mandatory = $true, HelpMessage = "Version of application")]
	[ValidateNotNullOrEmpty()]
	[string]$ApplicationVersion,
    
    [parameter(Mandatory = $true, HelpMessage = "SCCM Site Code Required")]
	[ValidateNotNullOrEmpty()]
	[string]$SiteCode,

	[parameter(Mandatory = $true, HelpMessage = "Base UNC path of your packages")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript({ Test-Path $_ })]
	[string]$RepositoryPath,

	[parameter(Mandatory = $true, HelpMessage = "MSI or another")]
	[ValidateNotNullOrEmpty()]
	[string]$DeploymentType


	
)
 
# Import SCCM PowerShell Module
$ModuleName = (get-item $env:SMS_ADMIN_UI_PATH).parent.FullName + "\ConfigurationManager.psd1"
Import-Module $ModuleName
 
Write-Debug "Site Code In Use : $SiteCode"

# Specify Adobe Flash Folder
# $AdobeFlashDir = "\Adobe Systems\Flash Player\"
 
# Remove back slash if added to the RepositoryPath
$RepositoryPath = $RepositoryPath.Trimend("\")
Write-Debug "Package UNC Base Path In Use : $RepositoryPath"
 
# Specify SCCM Collection - Temp
$TempDeviceCollection = "TempDeviceCollection"

# Specify SCCM Collection - Production
$ProductionDeviceCollection = "All Something"
 
$ApplicationName = $ApplicationName + ' ' + $ApplicationVersion

if ($DeploymentType -eq "MSI")
     
     {
 
        foreach ($MSI in (Get-ChildItem -Path ($RepositoryPath) -Filter *.MSI | Sort-Object $_.LastWriteTime))
        {
	        Set-Location -Path ($SiteCode + ":")
	
            if ((Get-CMApplication -Name $ApplicationName) -eq $true)
                {
	            Write-Debug "Creating application without optional icon specified"
		
		        # Create Application
		        New-CMApplication -Name "$ApplicationName" -LocalizedDescription $ApplicationName -LocalizedName $ApplicationName -SoftwareVersion $ApplicationVersion
		
		        # Create Deployment Types
		        Add-CMMsiDeploymentType -DeploymentTypeName $ApplicationName -ApplicationName $ApplicationName -MsiInstaller -ForceForUnknownPublisher $true -InstallationBehaviorType InstallForSystem -ins -Verbose
				
		        # Distripute Content
				Start-CMContentDistribution -ApplicationName $ApplicationName -DistributionPointName "srv-sccm-ps.rambler.ramblermedia.com" -Verbose
				
		        # Create Application Deployment
				Start-CMApplicationDeployment -CollectionName $TempDeviceCollection -Name $ApplicationName -AvailableDateTime (get-date) -DeployAction Install -DeployPurpose Required -TimeBaseOn LocalTime -OverrideServiceWindow $true -UserNotification HideAll
	
                # Refresh Policy on collection
                Start-Sleep -Seconds 10
 
                Invoke-CMClientNotification -DeviceCollectionName $TempDeviceCollection -NotificationType RequestMachinePolicyNow -Verbose

                # Run the Deployment Summarization
                Invoke-CMDeploymentSummarization -CollectionName $TempDeviceCollection -Verbose
	            }
    
            else { Write-Warning -Message "$ApplicationName already exists." }
        }
    }

            elseif ($ApplicationName -contains "Skype")
                 {
              Set-Location -Path ($SiteCode + ":")
                    
                if ((Get-CMApplication -Name $ApplicationName) -eq $true)
                    {
	
		        Write-Debug "Creating application without optional icon specified"
		       
                New-CMApplication -Name "$ApplicationName" -LocalizedDescription $ApplicationName -LocalizedName $ApplicationName -SoftwareVersion $ApplicationVersion
		
		        # Create Deployment Types
		        Add-CMMsiDeploymentType -DeploymentTypeName $ApplicationName -ApplicationName $ApplicationName -InstallationFileLocation $MSI.FullName -MsiInstaller -ForceForUnknownPublisher $true -InstallationBehaviorType InstallForSystem -InstallCommand "Deploy-Application.exe -DeploymentType 'Install' -DeployMode 'Silent'" -UninstallCommand "Deploy-Application.exe -DeploymentType 'Uninstall' -DeployMode 'Silent'" -Verbose
				
		        # Distripute Content
				
		        Start-CMContentDistribution -ApplicationName $ApplicationName -DistributionPointName "srv-sccm-ps.rambler.ramblermedia.com" -Verbose
				
		        # Create Application Deployment
				
		        Start-CMApplicationDeployment -CollectionName $TempDeviceCollection -Name $ApplicationName -AvailableDateTime (get-date) -DeployAction Install -DeployPurpose Required -TimeBaseOn LocalTime -OverrideServiceWindow $true -UserNotification HideAll
	
                # Refresh Policy on collection
 
                Start-Sleep -Seconds 10
 
                Invoke-CMClientNotification -DeviceCollectionName $TempDeviceCollection -NotificationType RequestMachinePolicyNow -Verbose

                # Run the Deployment Summarization
 
                Invoke-CMDeploymentSummarization -CollectionName $TempDeviceCollection -Verbose
	        }
    
            else
            {
            Write-Warning -Message "$ApplicationName already exists."
            }
              


                 }
 

