# Get Server IP using Powershell JOBs

begin
{
 $MaxThreads = 30 # Max concurrent running jobs.
 Get-Job | Remove-Job -Force
 
}
process
{


 $exec = {
  param([string] $s,[string] $sn )  
  process
  {
    try
    {
	   
	   $Timeout = 5000 #wait 5 secs for a reply
	   $Ping = New-Object System.Net.NetworkInformation.Ping
	   $Response = $Ping.Send($s,$Timeout)
	   if ($Response.Status -eq "Success")
	   {
		    $Obj = New-Object -TypeName PsObject -Property(@{
                'Server'  = $s;
                'IP_Address'= $Response.Address;
				        'Exception'= "N/A";
            })
            return $Obj
		}else
		{
			 $Response = $Ping.Send($sn,$Timeout)	
			 if ($Response.Status -eq "Success")
		   	 {
			    $Obj = New-Object -TypeName PsObject -Property(@{
	                'Server'  = $s;
	                'IP_Address'= $Response.Address;
					        'Exception'= "Retry Successful";
	            })
	            return $Obj
			 }
			 else
			 {
				 $Obj = New-Object -TypeName PsObject -Property(@{
		                'Server'  = $s;
		                'IP_Address'= "ERROR";
						        'Exception'= $Response.Status;
		         })
					return $Obj
			}
            
		}
    }catch
    {
       
		$Response = $Ping.Send($sn,$Timeout)	
		 if ($Response.Status -eq "Success")
	   	 {
		    $Obj = New-Object -TypeName PsObject -Property(@{
                'Server'  = $s;
                'IP_Address'= $Response.Address;
				        'Exception'= "Retry Successful";
            })
            return $Obj
		 }
		 else
		 {
		  $Obj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'IP_Address'= "ERROR";
						            'Exception'= $_.Exception.Message;
                     })
                     return $Obj
		}
    }
   }
}


   


     Function UpdateIP_Address([String] $CSHostname, [String]$ip_address,[String]$exception )
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: IpAddress: $ip_address, Exception: $exception" -logname "Get_ServerIP.log"

          $Query = "update [ServerAudit][ServerInfo] set [IP Address]='$ip_address' where CShostname='$CSHostname'" 
             
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "Get_ServerIP.log"
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
           foreach ($server in $HostNames)
           {
                $s = $server.CSHostname;
				$sn= $server.CSCIname;
                Start-Job -ScriptBlock $exec -ArgumentList $s,$sn -Name $s
                    
                            #Throttle jobs.
                            While($(Get-Job -State Running).Count -ge $MaxThreads) 
                            {
                                Get-Job | Wait-Job -Any -Timeout 120| Out-Null 
                            }
                            $progressCounter++
            }
            
          }
	      else
	        {
	            WriteLog -msg "List of Servers is empty" -logname "Get_ServerIP.log"
	        }
    }catch
    {
         $Obj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'IP_Address'= "ERROR";
						'Exception'= $_.Exception.Message;
                     })

                     $insert = UpdateIP_Address -CSHostname $Obj.Server -ip_address $Obj.IP_Address -exception $Obj.Exception
    }
    Get-Job | Wait-Job -Timeout 180
    $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName 
     foreach ($r in $results){
	 	$update = UpdateIP_Address -CSHostname $r.Server -ip_address $r.IP_Address -exception $r.Exception
	 }

     #Sending E-mail
     SendEmail -logname "Get_ServerIP.log"
     WriteLog -msg "Process has Finished." -logname "Get_ServerIP.log"
}