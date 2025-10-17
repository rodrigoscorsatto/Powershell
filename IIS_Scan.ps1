<#
Check if IIS service is available in the remote servers using powershell jobs
#>

param (

    [Parameter(Mandatory=$false,ValueFromPipeline=$true,HelpMessage='Please inform the file path containing the Servers list.')][string] $HostNameList
 
);

begin
{
    $MaxThreads = 30 # Max concurrent running jobs.
    Write-Host "Script started at" (Get-Date).ToString()
    Get-Job | Remove-Job -Force
   


}
process
{
$exec = {
  param([string] $server)  
  process{
   $ServerInfo = @()
   $version = $null 
   try
   { 

    $iis = get-service w3svc -ComputerName $server -ErrorAction SilentlyContinue
    if ($iis.Status -eq "Running") 
    {

        $reg = Get-WmiObject -List StdRegProv -ComputerName $server -Namespace root\default -ErrorAction SilentlyContinue
        if ($reg) 
        {
		    $HKLM = 2147483650
		    $regVer = $reg.GetStringValue($HKLM,"SOFTWARE\Microsoft\InetStp","VersionString").sValue
		    if ($regVer)
            {
		      $version= $regVer.replace('Version','').trim()
			}
			else
            {
                 $version= "No IIS identified" 
            }
		}
		else
        {   
           $version= "ERROR - Could not collect $server information from registry"
            
        }
    
        $ServerInfo = New-Object -TypeName PsObject -Property(@{
                'Server' = [string] $server;
                'IIS'="IIS";
                'Version' = [string] $version;
            })
            return $ServerInfo

    } 
    else {
        
     $ServerInfo = New-Object -TypeName PsObject -Property(@{
                'Server' = [string] $server;
                'IIS'="Custom Service";
                'Version' = "N/A";
            })
            return $ServerInfo

    }
     

            }catch{
                
                $ServerInfo = New-Object -TypeName PsObject -Property(@{
                'Server' = [string] $server;
                'IIS'="ERROR: $_.Exception.Message";
                'Version' = "ERROR";
            })  
            }
            return $ServerInfo
  }
}

 



     Function UpdateIISData([String] $Server, [String]$AppType, [String] $version)
     {
       try{ 

          WriteLog -msg "Updating Server: $Server with AppType: $AppType IIS Version: $version" -logname "IIS_scan.log"

          $Query = "UPDATE Server SET AppType='$AppType', IIS_Version='$version' WHERE CSHostname='$Server'"
          
          $conn = New-Object System.Data.SqlClient.SqlConnection
          
          $connectionString = GetConnectionString -DatabaseName "ServerAudit"
		  $conn.ConnectionString = $connectionString
          $conn.open()
          $cmd = New-Object System.Data.SqlClient.SqlCommand
          $cmd.connection = $conn

          $cmd.commandtext = $Query
          $result = $cmd.executenonquery()

       }catch
          {
            WriteLog -msg "There is an error when trying to update Data. Error: $_.Exception.Message" -logname "IIS_scan.log"  
          }
           finally
           {
               $conn.close()

           }
     }

     

     Function GetServerListError()
     {
       $table = $null 
       try{
              			  
              $connectionString = GetConnectionString -DatabaseName "ServerAudit"
			  $connection = New-Object System.Data.SqlClient.SqlConnection
              $Query = GetServerErrorListIISScanQuery
              $connection.ConnectionString = $connectionString
              $connection.Open()
              $command = $connection.CreateCommand()
              $command.CommandText  = $query
              $result = $command.ExecuteReader()
              $table = new-object "System.Data.DataTable"
              $table.Load($result)
               
              return $table
          }catch
          {
            WriteLog -msg "There is an error when trying to get Server Data. Error: $_.Exception.Message" -logname "IIS_scan.log"           
            $table=$null
          }
          finally
          {
            $connection.close()

          } 
     }

Function RetrywithInvoke([String] $server)
{
   try
   { 

        WriteLog "Retrying Server: $server" -logname "IIS_scan.log"
        $reg = Invoke-Command -ComputerName $server -ScriptBlock {Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\InetStp' -Name "VersionString" |Select-Object VersionString} -Credential $ServerCredential
        if ($reg) 
        {
             $version= $reg.VersionString
        }else
        {
            $version= "ERROR - Could not collect version information from registry"
        }    
        $ServerInfo = New-Object -TypeName PsObject -Property(@{
                'Server' = [string] $server;
                'IIS'="IIS";
                'Version' = [string] $version;
            })
            return $ServerInfo

    }catch
     {
                
          $ServerInfo = New-Object -TypeName PsObject -Property(@{
             'Server' = [string] $server;
             'IIS'="ERROR";
             'Version' = "ERROR: $_.Exception.Message";
          })  
      }
            return $ServerInfo
}

     
            
        if ($HostNameList)
        {
            $HostNames = Get-Content $HostNameList
            $progressCounter=0
            foreach ($s in $HostNames)
            {
                    $progressPercentage = [int]([float]($progressCounter / $HostNames.Count) * 100)
                    Write-Progress -Activity "Servers" -CurrentOperation "Processing server $s" -PercentComplete $progressPercentage -Status "$progressPercentage% Complete, Server Count: $progressCounter"
                    Start-Job -ScriptBlock $exec -ArgumentList $s, $ServerCredential -Name $s -Credential $ServerCredential 
                    
                    
                    #Throttle jobs.
                    While($(Get-Job -State Running).Count -ge $MaxThreads) 
                    {
                        #Write-host "Waiting Threads to release..." -ForegroundColor Yellow
                        Get-Job | Wait-Job -Any -Timeout 180| Out-Null 
                    }
                    $progressCounter++
               }
        }else
        {
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
                        #Write-host "Waiting Threads to release..." -ForegroundColor Yellow
                        Get-Job | Wait-Job -Any -Timeout 180| Out-Null 
                    }
                    $progressCounter++
               }
            }else
            {
                Write-Host ("There is an error when trying to get Server List from database")
            }
        }
        Get-Job | Wait-Job -Timeout 180
        $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName     

        foreach ($result in $results)
        {
            UpdateIISData -Server $result.Server -AppType $result.IIS -version $result.Version
        }

        # retying servers with errors.
        $ErrorHostNames = GetServerListError
        foreach ($server in $ErrorHostNames)
        {
            $server=$server.CSHostname
            $result = RetrywithInvoke -server $server
            if ($result) 
            {
                UpdateIISData -Server $result.Server -AppType $result.IIS -version $result.Version
            }

        }
       
        #Sending E-mail
        SendEmail -logname "IIS_scan.log"
        WriteLog -msg "Process has Finished." -logname "IIS_scan.log"
    }