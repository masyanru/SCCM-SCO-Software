
# common variables section



$softwareCSV = "C:\runbooks\software automation\1. Check Versions\versions.csv"
$tempDownloadFolder = "C:\runbooks\software automation\1. Check Versions\download\temp"
$downloadFolder = "C:\runbooks\software automation\1. Check Versions\download"
$runbookPath = "C:\runbooks\software automation\1. Check Versions"




  function software:adobe {
  
  $releases = "https://get.adobe.com/en/flashplayer/" # URL to for GetLatest

  # 
  $HTML = Invoke-WebRequest -Uri $releases
  $try = ($HTML.ParsedHtml.getElementsByTagName('p') | Where{ $_.className -eq 'NoBottomMargin' } ).innerText
  $try = $try  -split "\r?\n"
  $try = $try[0] -replace ' ', ' = '
  $try =  ConvertFrom-StringData -StringData $try
  $CurrentVersion = ( $try.Version )
  $majorVersion = ([version] $CurrentVersion).Major

  # Flash Player Active X
  $softwareAdobeUrl = @()
  $softwareAdobeUrl += "https://download.macromedia.com/pub/flashplayer/pdc/${CurrentVersion}/install_flash_player_${majorVersion}_active_x.msi"
  $softwareAdobeUrl += "https://download.macromedia.com/get/flashplayer/pdc/${CurrentVersion}/install_flash_player_${majorVersion}_plugin.msi"
  $softwareAdobeUrl += "https://download.macromedia.com/pub/flashplayer/pdc/${CurrentVersion}/install_flash_player_${majorVersion}_ppapi.msi"
  
  $CurrentVersion
  $softwareAdobeUrl

  $csv = Import-Csv $softwareCSV -delimiter ',' -Encoding default

  $csv | % {

              if ($_.Name -eq "adobe_flash" -and $_.Old -lt $CurrentVersion)
                  { 
                                Write-Host $true
                                $_.New = $CurrentVersion 
                    
                                 # create folder, download msi
                                 New-Item -ItemType Directory -Path "$downloadFolder\adobe $CurrentVersion\" -Force
                                 $targetDir = "$downloadFolder\adobe $CurrentVersion\"            
                                 function DownloadFile([Object[]] $sourceFiles,[string]$targetDirectory) {            
                                 $wc = New-Object System.Net.WebClient            
                                 $sourceFile = $softwareAdobeUrl

                                     foreach ($sourceFile in $sourceFile){            
                                      $sourceFileName = $sourceFile.SubString($sourceFile.LastIndexOf('/')+1)            
                                      $targetFileName = $targetDirectory + $sourceFileName            
                                      $wc.DownloadFile($sourceFile, $targetFileName)            
                                      Write-Host "Downloaded $sourceFile to file location $targetFileName"             
                                     }            
            
                                }            
            
                                DownloadFile $softwareAdobeUrl $targetDir   
                     
                  }
               
          $csv | Export-Csv $softwareCSV -NoTypeInformation

                 }


                       }



  function software:skype {

  $skypeURL = "http://www.skype.com/go/getskype-msi" 
  $output = "$tempDownloadFolder\skype.msi"
  
  # download skype msi
  Import-Module BitsTransfer
  Start-BitsTransfer -Source $skypeURL -Destination $output
   
  # get msi version
  $skypeCurrentVersion = Get-MsiDatabaseVersion "$tempDownloadFolder\skype.msi"
  $skypeCurrentVersion = [string]$skypeCurrentVersion -replace '\s',''

  New-Item -ItemType Directory -Path "$downloadFolder\skype $skypeCurrentVersion\" -Force
  # Wait-FileUnlock "C:\runbooks\software automation\1. Check Versions\download\skype $skypeCurrentVersion\" -v
  Copy-Item "$tempDownloadFolder\skype.msi" "$downloadFolder\skype $skypeCurrentVersion\" -Force
    
  # read CSV
  $csv = Import-Csv $softwareCSV -delimiter ',' -Encoding default
 
    foreach ($row in $csv) 
            {
             if ($row.Name -eq "skype" -and [System.Version]$row.Old -lt [System.Version]$skypeCurrentVersion)
                 { 
                   Write-Host $true + $row.Name
                   $row.New = $skypeCurrentVersion -join '.'
                          
                 }
              else
                 {
                  Write-Host $false
                 }
            }
     $csv | Export-Csv $softwareCSV -NoTypeInformation
            }


  function Get-MsiDatabaseVersion {
    param (
        [IO.FileInfo] $FilePath
    )

    try {
        $windowsInstaller = New-Object -com WindowsInstaller.Installer

        $database = $windowsInstaller.GetType().InvokeMember(
                "OpenDatabase", "InvokeMethod", $Null, 
                $windowsInstaller, @($FilePath.FullName, 0)
            )

        $q = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
        $View = $database.GetType().InvokeMember(
                "OpenView", "InvokeMethod", $Null, $database, ($q)
            )

        $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null)

        $record = $View.GetType().InvokeMember(
                "Fetch", "InvokeMethod", $Null, $View, $Null
            )

        $productVersion = $record.GetType().InvokeMember(
                "StringData", "GetProperty", $Null, $record, 1
            )

        $View.GetType().InvokeMember("Close", "InvokeMethod", $Null, $View, $Null)

        return $productVersion
       

    } catch {
        throw "Failed to get MSI file version the error was: {0}." -f $_
    }


}





Function Wait-FileUnlock{
    Param(
        [Parameter()]
        [IO.FileInfo]$File,
        [int]$SleepInterval=500
    )
    while(1){
        try{
           $fs=$file.Open('open','read', 'Read')
           $fs.Close()
            Write-Verbose "$file not open"
           return
           }
        catch{
           Start-Sleep -Milliseconds $SleepInterval
           Write-Verbose '-'
        }
	}
}




# function to create application in SCCM

function Create-SCCMApplication{
    Param(
    [string]$ApplicatioName
    
    )







}
