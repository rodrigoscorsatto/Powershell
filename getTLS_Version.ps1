# Get TLS version from the remote servers using powershell jobs
begin
{
 $MaxThreads = 30 # Max concurrent running jobs.
 Get-Job | Remove-Job -Force
}
process
{


 $exec = {
  param([string] $s)  
  process
  {
    try
    {
      $reg = e:\get-registry.ps1 -ComputerName $s -KeyPath "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols" -HKLM -Recursive -ValueName "Enabled"
      
		$TLSObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'TLS_List'= $reg;
						'Exception'  = "N/A";
                     })
					 return $TLSObj
    }catch
    {
       $TLSObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
						'TLS_List'= "ERROR"
                        'Type'= "ERROR";
                        'Protocol'    = "ERROR";
                        'Status'= "ERROR";
                        'Exception'  = "$_.Exception.Message";
                     })
                     return $TLSObj

    }



  }
}

   
    
     Function InsertTLSScanData([String] $CSHostname, [String]$Type, [String]$Protocol, [String]$Status, [String]$Exception)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Type: $Type, Protocol: $Protocol, Status: $Status, Exception: $Exception" -logname "TLS_Scan.log" 
          $Query =@"
          INSERT INTO TLSScan (CSHostname, Type, Protocol,[Status],Exception) VALUES('$CSHostname','$Type','$Protocol','$Status','$Exception');
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "TLS_Scan.log" 
            $result=$null
          }
          finally
          {
            $conn.close()

          }
     }


try{
    
        $HostNames = GeServerListAll
        
         if ($HostNames -ne $null)
         { 
           $delete = DeleteAllDataTable -tablename "TLSScan" -databasename "ServerAudit" -logname "TLS_Scan.log"
           for ($i=0; $i -lt $HostNames.Length; $i++)
           {
                $s = $HostNames[$i].CSHostname;
                
                if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ErrorAction SilentlyContinue -quiet)
                {
                    if (Test-WSMan -ComputerName $s -ErrorAction SilentlyContinue) 
                    {
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

                     }else
                     {
                        $TLSObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'Type'= "ERROR";
                        'Protocol'    = "ERROR";
                        'Status'= "ERROR";
                        'Exception'  = "ERROR: The client cannot connect to the destination specified in the request";
                     })
                     $insert = InsertTLSScanData -CSHostname $TLSObj.Server -Type $TLSObj.Type -Protocol $TLSObj.Protocol -Status $TLSObj.Status -Exception $TLSObj.Exception
                     
                     }
                }else
                {
				
                 $TLSObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'Type'= "ERROR";
                        'Protocol'    = "ERROR";
                        'Status'= "ERROR";
                        'Exception'  = "ERROR: Server is not Reachable";
                     })
                     $insert = InsertTLSScanData -CSHostname $TLSObj.Server -Type $TLSObj.Type -Protocol $TLSObj.Protocol -Status $TLSObj.Status -Exception $TLSObj.Exception
                }
            }
            
        }
        else
        {
            WriteLog -msg "List of Servers is empty" -logname "TLS_Scan.log" 
        }
    }catch
    {
		$TLSObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'Type'= "ERROR";
                        'Protocol'    = "ERROR";
                        'Status'= "ERROR";
                        'Exception'  = "$_.Exception.Message";
                     })
                     $insert = InsertTLSScanData -CSHostname $TLSObj.Server -Type $TLSObj.Type -Protocol $TLSObj.Protocol -Status $TLSObj.Status -Exception $TLSObj.Exception
        
    }
    Get-Job | Wait-Job -Timeout 180
    $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName 
    
    foreach ($r in $results){

        $TLSList = $r.TLS_List
    
        if ($TLSList -eq "ERROR")
        {
            $insert = InsertTLSScanData -CSHostname $r.Server -Type $r.Type -Protocol $r.Protocol -Status $r.Status -Exception $r.Exception

        }else
        {

            foreach ($TLS in $TLSList)
            {
              $Protocol = $TLS.ParentKeyName
              $Type = $TLS.KeyName # test for Key Enabled
              if ($TLS.Value -ne "0")
              {
                $Status="Enabled"
              }
              else
              {
              $Status = "Disabled"
              }			
              $insert = InsertTLSScanData -CSHostname $r.Server -Type $Type -Protocol $Protocol -Status $Status -Exception $r.Exception
            }
        }
    }
    #Sending E-mail
    SendEmail -logname "TLS_Scan.log" 
    WriteLog -msg "Process has Finished." -logname "TLS_Scan.log" 
}