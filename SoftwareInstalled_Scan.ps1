<#
.DESCRIPTION
    1)	You can use this script to scan the servers and get list of software installed.


.NOTES
    File Name      : Scan_SoftwareInstalled.ps1
    Author         : Rodrigo Scorsatto 
    Prerequisite   : PowerShell V2 or upper
    Version        : 1.0 - 06/21/2021

#>


begin
{
 
 $MaxThreads = 30 # Max concurrent running jobs.
 Get-Job | Remove-Job -Force
 
}
process
{
$exec = {
  param([string] $serverName)  
  process{
   
        try
        {
            $softwareList = Get-WmiObject -Class Win32_Product  -ComputerName $serverName  | select-object Name, Vendor, Version, Caption, InstallLocation, InstallSource, InstallDate

            $result = @()
		    foreach ($item in $softwareList)
            {
                $date=$item.InstallDate;
                if ($date)
                {
		            $installdate = [datetime]::ParseExact($date , "yyyyMMdd",$null)
                }else
                {
                   $installdate = [datetime]"1999-01-01 00:00"; 
                }
                $Obj = New-Object -TypeName PsObject -Property(@{
                    'CSHostname'  = $serverName;
                    'Name'= [String]$item.Name;
                    'Vendor'=[String] $item.Vendor;
                    'Version'= [String]$item.Version;
                    'Caption'=[String]$item.Caption;
                    'InstallLocation'=[String]$item.InstallLocation;
                    'InstallSource'=[String]$item.InstallSource;
                    'InstallDate' = $installdate;
                    'Exception'= "N/A";
                })
                $result += $Obj
            }

            return $result
        }catch
        {
               $Exception = $_.Exception.Message;
               $Obj = New-Object -TypeName PsObject -Property(@{
                    'CSHostname'  = $serverName;
                    'Name'= "Error";
                    'Vendor'="Error";
                    'Version'= "Error";
                    'Caption'="Error";
                    'InstallLocation'="Error";
                    'InstallSource'="Error";
                    'InstallDate' = [datetime]"1999-01-01 00:00";
                    'Exception'= $Exception;
                })
                $result += $Obj
                return $result
        }
   }
}


   

     Function InsertSoftwareInstalledData([String] $CSHostname, [String]$Name, [String]$Vendor, [String]$Version, [String]$Caption, [String]$InstallLocation, [String]$InstallSource,[datetime]$InstallDate,[String]$Exception)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Name: $Name, Vendor: $Vendor,Version: $Version, Caption: $Caption, Install Location: $InstallLocation, Install Source: $InstallSource, InstallDate: $installdate Exception: $Exception" -logname "SoftwareInstalledScan.log"

          $Query = "INSERT INTO Software_Installed_Scan (CSHostname, Name,Vendor,Version,Caption,InstallLocation,InstallSource,InstallDate, Exception) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')" -f $CSHostname,$Name,$Vendor,$Version,$Caption,$InstallLocation,$InstallSource,$installdate,$Exception
          $conn = New-Object System.Data.SqlClient.SqlConnection

          $connectionString = GetConnectionString -DatabaseName "ServerAudit"
		  $conn.ConnectionString = $connectionString
          $conn.open()
          $cmd = New-Object System.Data.SqlClient.SqlCommand
          $cmd.connection = $conn

          $cmd.commandtext = $Query
          $result = $cmd.executenonquery()
          return $result
       }catch
          {
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "SoftwareInstalledScan.log"
            $result=$null
          }
          finally
          {
            $conn.close()

          }
     }

    
$HostNames = GeServerListAll
   
if ($HostNames -ne $null)
{ 
    $progressCounter=0
    for ($i=0; $i -lt $HostNames.Length; $i++)
    {
        $s = $HostNames[$i].CSHostname;
        $progressPercentage = [int]([float]($progressCounter / $HostNames.Length) * 100)
        Write-Progress -Activity "Servers" -CurrentOperation "Processing server $s" -PercentComplete $progressPercentage -Status "$progressPercentage% Complete, Server Count: $progressCounter"
        Start-Job -ScriptBlock $exec -ArgumentList $s -Name $s               
                    
        #Throttle jobs.
        While($(Get-Job -State Running).Count -ge $MaxThreads) 
        {
            Get-Job | Wait-Job -Any -Timeout 180| Out-Null 
        }
        $progressCounter++
    }
}else
    {
        WriteLog -msg "There is an error when trying to get Server List from database" -logname "SoftwareInstalledScan.log"
    }
        
        Get-Job | Wait-Job -Timeout 180
        $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName 

        $delete =  DeleteAllDataTable -tablename "Software_Installed_Scan" -databasename "ServerAudit" -logname "SoftwareInstalledScan.log"
        foreach ($result in $results)
        {
            $insert = InsertSoftwareInstalledData -CSHostname $result.CSHostname -Name $result.Name -Vendor $result.Vendor -Version $result.Version -Caption $result.Caption -InstallLocation $result.InstallLocation -InstallSource $result.InstallSource -InstallDate $result.InstallDate -Exception $result.Exception
        }
        
        #Sending E-mail
        SendEmail -logname "SoftwareInstalledScan.log"
        WriteLog -msg "Process has Finished." -logname "SoftwareInstalledScan.log"
    }