
<#
.DESCRIPTION
    1)	You can use this script set custom headers to IIS. This script will change it on ALL sites. the results will be exported to a CSV file


.NOTES
    File Name      : SetCustomHeaderToIIS.ps1
    Author         : Rodrigo Scorsatto
    Prerequisite   : PowerShell V2 or upper, Windows 2012 or above
    Version        : 1.0 - 04/30/2019

#>

param (

    [Parameter(Mandatory=$false,ValueFromPipeline=$true,HelpMessage='Please inform the AppPool name.')][string] $UserName,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage='Please inform the file path containing the Servers list.')][string] $HostNameList,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage='Please inform the Header Name to be included.')][string] $HeaderList
  
);

$ToExecute = 
{
     param([string] $server,  $HList )  
    
  try
  {
	Import-Module WebAdministration;
    Add-Type -AssemblyName "Microsoft.Web.Administration, Version=7.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35, processorArchitecture=MSIL"

    #$iisObj = new-object Microsoft.Web.Administration.ServerManager 
    foreach ($webapp in Get-ChildItem -Path IIS:\Sites)
    {
        $iisObj = new-object Microsoft.Web.Administration.ServerManager
        foreach ($header in $HList)
        {
            $header | %{ $header = $_ -split ':'; $HeaderName=$header[0]; $HeaderValue=$header[1] }
            $WebSiteName = $webapp.name
            Write-Host "Http Response Headers: Adding Header Name: $HeaderName for Site Name: $WebSiteName"  -ForegroundColor Yellow           

            #$PSPath =  'MACHINE/WEBROOT/APPHOST/' + $WebSiteName          
            #Remove-WebConfigurationProperty -PSPath $PSPath -Name . -Filter system.webServer/httpProtocol/customHeaders -AtElement @{name =$HeaderName }            

            $config = $iisObj.GetWebConfiguration($WebSiteName)
            $httpProtocolSection = $config.GetSection("system.webServer/httpProtocol")
            $customHeadersCollection = $httpProtocolSection.GetCollection("customHeaders")
            $headerFound = $false
            foreach ($header in $customHeadersCollection)
            {
              if ($HeaderName -eq $header.GetAttributeValue("name"))
              {
               $headerFound = $true
              }
              
            }
            if ($headerFound)
            {
                Write-Host "Header $HeaderName already Exists. No Action Performed." -ForegroundColor Yellow
            }else
            {
                $addElement = $customHeadersCollection.CreateElement("add")
                $addElement["name"] = $HeaderName
                $addElement["value"] = $HeaderValue
                $op = $customHeadersCollection.Add($addElement)
            }
            
        }
        $changes = $iisObj.CommitChanges()
    }

    foreach ($webapp in Get-ChildItem -Path IIS:\Sites)
    {

        $os = (Get-WmiObject Win32_OperatingSystem).name
        if (($os -like "*2008*") -or ($os -like "*2012 Datacenter*"))
        {
            Write-Host "OS Version: $os is not supported to add Custom Fields." -ForegroundColor Red
        }
        else
        {
            foreach ($header in $HList)
            {
               $header | %{ $header = $_ -split ':'; $HeaderName=$header[0]; $HeaderValue=$header[1] }
               $WebSiteName = $webapp.name
               Write-Host "Logging: Adding Custom Fields: $HeaderName for Site Name: $WebSiteName" -ForegroundColor Yellow
               New-ItemProperty "IIS:\Sites\$WebSiteName" -Name logfile.customFields.collection -Value @{logFieldName=$HeaderName;sourceType='RequestHeader';sourceName=$HeaderName}
            }
        }
    }

    $HeaderObj = New-Object -TypeName PsObject -Property(@{
            'Server'  = $server;
            'Status'  = "Success";
        })
    Write-Host "Header added Successfully to All Sites!" -ForegroundColor Green
    return $HeaderObj
  }
  Catch 
  {
      $HeaderObj = New-Object -TypeName PsObject -Property(@{
            'Server'  = $server;
            'Status'  = "ERROR: $_.Exception.Message";
        })
    write-host "Error, Unable to add Custom Header to All Sites, Operation Aborted! Error Info:" $_.Exception.Message -ForegroundColor Red
    return $HeaderObj
  }  
}


$date = Get-Date -f MM-dd-yyyy-HHmm  
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($HostNameList)
$filename = $date+"_"+$baseName+"_"+"SetCustomHeaderToIIS.csv"
$path = ".\" + $filename
Start-Transcript ./SetcustomHeaderToIIS.log

if ( $UserName -eq $Cred.UserName )
{
    Write-Host "Using ( $Cred.UserName )"
}
else
{
    $Cred = Get-Credential
}

    if (Test-Path $HostNameList)
    {
        $HostNames = Get-Content $HostNameList
        
        if (Test-Path $HeaderList){
             if ($HostNames.Length -gt 0 ) 
            {
                $HList = Get-Content $HeaderList
                foreach ($s in $HostNames)
                {
                    if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet)
                    {
                    Write-Host "Processing Server: "$s
                    $res = Invoke-command -computer $s -ScriptBlock $ToExecute -ArgumentList $s,$HList -credential $Cred -SessionOption (New-PSSessionOption -OperationTimeout 90000)
                    $res | Select-Object Server, Status | Export-Csv $path  -Append -NoTypeInformation
                    }else
                    {
                       $servercheck = New-Object -TypeName PsObject -Property(@{
                            'Server'  = $s;
                            'Status'  = "ERROR: Server is not Reacheable";
                       })
                       $servercheck | Select-Object Server, Status | Export-Csv $path  -Append -NoTypeInformation
                    }
                }
            }
            else
            {
                Write-Error ("List Servers empty")
            }
        }else
        {
         Write-Error ("No valid Header list file: $HeaderList")
        }   
 
    }
    else
    {
        Write-Error ("No valid Server list file: $HostNameList")
    }
    Stop-Transcript
    Write-Host (" Script has finished!") -ForegroundColor Green