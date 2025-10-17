# Script to scan for services in the remote servers
begin
{
 $MaxThreads = 30 # Max concurrent running jobs.
 Get-Job | Remove-Job -Force
 
}
process
{


 $exec = {
  param([string] $s, $blackList)  
  process
  {
    try
    {
      $service = Get-WmiObject -Query "select * from Win32_Service" -ComputerName $s -ErrorAction Stop
      $ServiceList = $service | where {$_.caption -notin $blackList}
	  if ($ServiceList -eq $null)	
	  {
	    $ServiceObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'ServiceName'= "No Custom Service Found";
                        'DisplayName'= "No Custom Service Found";
                        'StartupType'= "No Custom Service Found";
                        'Status'= "No Custom Service Found";
                        'Description'= "No Custom Service Found";
                        'Path' = "No Custom Service Found";
                        'logon'= "No Custom Service Found";
                     })
                     return $ServiceObj
 
	  }else
	  {
       $ServiceObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'ServiceName'= $ServiceList;
                     })
                     return $ServiceObj

      }     

    }catch
    {
        $ServiceObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'ServiceName'= "ERROR";
                        'DisplayName'    = "ERROR";
                        'StartupType'= "ERROR";
                        'Status'  = "ERROR";
                        'Description' = "ERROR: $_.Exception.Message";
                        'Path' = "ERROR";
                        'logon'       = "ERROR";
                     })
                     return $ServiceObj

    }



  }
}

    
     Function GetServiceBlackList()
     {
       $table = $null 
       try{		  
              $connectionString = GetConnectionString -DatabaseName "ServiceManagement"
			        $connection = New-Object System.Data.SqlClient.SqlConnection
              $Query = GetServiceBlackListQuery
              $connection.ConnectionString = $connectionString
              $connection.Open()
              $command = $connection.CreateCommand()
              $command.CommandText  = $query
              $result = $command.ExecuteReader()
              $table = new-object “System.Data.DataTable”
              $table.Load($result)
              return $table
          }catch
          {
            WriteLog -msg "There is an error when trying to get server data. Error: $_.Exception.Message" -logname "ServiceScan.log"
            $table=$null
          }
          finally
          {
            $connection.close()

          }

     }

     Function InsertServiceScanData([String] $CSHostname, [String]$ServiceName, [String]$DisplayName, [String]$Description, [String]$StartupType, [String]$Status, [String]$Path, [String] $Logon)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: ServiceName: $ServiceName, StartupType: $StartupType, Status: $Status" -logname "ServiceScan.log"

          $Description = $Description -replace("'",""); 
          $Query =@"
          INSERT INTO ServiceScan (CSHostname, ServiceName, DisplayName,StartupType,[Status],[Description],[Path], logon) VALUES('$CSHostname','$ServiceName','$DisplayName','$StartupType','$Status','$Description','$Path','$Logon');
"@
           
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "ServiceScan.log"
            $result=$null
          }
          finally
          {
            $conn.close()

          }
     }


try{
    
        $HostNames = GeServerListAll
        $BL = GetServiceBlackList
        $BlackList = $BL|%{ [string]$($_.serviceName)}
        
         if ($HostNames -ne $null)
         { 
           $delete = DeleteAllDataTable -tablename "ServiceScan" -databasename "ServerAudit" -logname "ServiceScan.log"
           for ($i=0; $i -lt $HostNames.Length; $i++)
           {
                $s = $HostNames[$i].CSHostname;
                
                if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ErrorAction SilentlyContinue -quiet)
                {
                    if (Test-WSMan -ComputerName $s -ErrorAction SilentlyContinue) 
                    {
                            $progressPercentage = [int]([float]($progressCounter / $HostNames.Length) * 100)
                            Write-Progress -Activity "Servers" -CurrentOperation "Processing server $s" -PercentComplete $progressPercentage -Status "$progressPercentage% Complete, Server Count: $progressCounter"
                            Start-Job -ScriptBlock $exec -ArgumentList $s, $BlackList -Name $s
                    
                            #Throttle jobs.
                            While($(Get-Job -State Running).Count -ge $MaxThreads) 
                            {
                                #Write-host "Waiting Threads to release..." -ForegroundColor Yellow
                                Get-Job | Wait-Job -Any -Timeout 180| Out-Null 
                            }
                            $progressCounter++

                     }else
                     {
                        $ServiceObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'ServiceName'= "ERROR";
                        'DisplayName'    = "ERROR";
                        'StartupType'= "ERROR";
                        'Status'  = "ERROR";
                        'Description' = "ERROR: The client cannot connect to the destination specified in the request";
                        'Path' = "ERROR";
                        'logon'       = "ERROR";
                     })
                     $insert = InsertServiceScanData -CSHostname $ServiceObj.Server -ServiceName $ServiceObj.ServiceName -DisplayName $ServiceObj.DisplayName -Description $ServiceObj.description -StartupType $ServiceObj.StartupType -Status $ServiceObj.Status -Path $ServiceObj.Path -Logon $ServiceObj.logon
                     
                     }
                }else
                {
                   $ServiceObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'ServiceName'= "ERROR";
                        'DisplayName'    = "ERROR";
                        'StartupType'= "ERROR";
                        'Status'  = "ERROR";
                        'Description' = "ERROR: Server is not Reachable";
                        'Path' = "ERROR";
                        'logon'       = "ERROR";
                     })
                     $insert = InsertServiceScanData -CSHostname $ServiceObj.Server -ServiceName $ServiceObj.ServiceName -DisplayName $ServiceObj.DisplayName -Description $ServiceObj.description -StartupType $ServiceObj.StartupType -Status $ServiceObj.Status -Path $ServiceObj.Path -Logon $ServiceObj.logon
                }
            }
            
        }
        else
        {
            WriteLog -msg "List of Servers is empty" -logname "ServiceScan.log"
        }
    }catch
    {
        $ServiceObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'ServiceName'= "ERROR";
                        'DisplayName'    = "ERROR";
                        'StartupType'= "ERROR";
                        'Status'  = "ERROR";
                        'Description' = "ERROR: $_.Exception.Message";
                        'Path' = "ERROR";
                        'logon'       = "ERROR";
                     })
                     $insert = InsertServiceScanData -CSHostname $ServiceObj.Server -ServiceName $ServiceObj.ServiceName -DisplayName $ServiceObj.DisplayName -Description $ServiceObj.description -StartupType $ServiceObj.StartupType -Status $ServiceObj.Status -Path $ServiceObj.Path -Logon $ServiceObj.logon
    }
    Get-Job | Wait-Job -Timeout 180
    $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName 
    
    foreach ($r in $results){

        $ServerList = $r.ServiceName
    
        if ($ServerList -eq "ERROR" -or $ServerList -eq "No Custom Service Found")
        {
            $insert = InsertServiceScanData -CSHostname $r.Server -ServiceName $r.ServiceName -DisplayName $r.DisplayName -Description $r.description -StartupType $r.StartupType -Status $r.Status -Path $r.Path -Logon $r.logon

        }else
        {

            foreach ($service in $ServerList)
            {
                 $insert = InsertServiceScanData -CSHostname $r.Server -ServiceName $service.Name -DisplayName $service.Caption -Description $service.description -StartupType $service.Startmode -Status $service.State -Path $service.pathname -Logon $service.StartName
            }
        }
    }
     #Sending E-mail
        SendEmail -logname "ServiceScan.log"
        WriteLog -msg "Process has Finished." -logname "ServiceScan.log"
}