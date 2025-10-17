<#
.DESCRIPTION
    1)	You can use this script to Enable App Pool Recycle Logging for each AppPool on Remote Server


.NOTES
    File Name      : Set_IIS_PoolRecycling_Logging.ps1
    Author         : Rodrigo Scorsatto 
    Prerequisite   : PowerShell V2 or upper
    Version        : 1.0 - 10/07/2019

#>

$ToExecute = 
{
     param([string] $server, $PoolName)  
  try
  {  
    Import-Module WebAdministration
    Set-WebConfigurationProperty -pspath 'MACHINE/WEBROOT/APPHOST'  -filter "system.applicationHost/applicationPools/add[@name='$PoolName']/recycling" -name "logEventOnRecycle" -value "Time,Requests,Schedule,Memory,IsapiUnhealthy,OnDemand,ConfigChange,PrivateMemory" 
    
    $ReturnObj = New-Object -TypeName PsObject -Property(@{
            'Server'  = $server;
			'PoolName'  = $PoolName;
			'Status'  = "Success";
			'Exception'= "N/A";
        })
    return $ReturnObj
  }
  catch
  {
         $ReturnObj = New-Object -TypeName PsObject -Property(@{
            'Server'  = $server;
			'PoolName'  = $PoolName;
			'Status'  = "ERROR";
			'Exception'= "ERROR: $_.Exception.Message";
        })
        return $ReturnObj
  }
}


	
  Function InsertServerIISAppPoolRecycleLoggingFix([String] $CSHostname, [String]$PoolName, [String]$Status, [String]$Exception)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Pool:$PoolName, Status: $Status, Exception: $Exception"  
		      $Exception = $Exception -replace("'",""); 
          $Query = "INSERT INTO ServerIISAppPoolRecycleLoggingFix (CSHostname, PoolName, Status, Exception) VALUES('{0}','{1}','{2}','{3}')" -f $CSHostname,$PoolName,$Status, $Exception
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "ServerIISAppPoolRecycleLoggingFix.log"
            $result=$null
          }
          finally
          {
            $conn.close()

          }
     }

     Function GetServerList()
     {
       $table = $null 
       try{
              
              $connectionString = GetConnectionString -DatabaseName "ServerAudit"
			        $connection = New-Object System.Data.SqlClient.SqlConnection
              $Query = GetServerListIISAppPoolLoggingFixQuery

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
            WriteLog -msg "There is an error when trying to get server data. Error: $_.Exception.Message" -logname "ServerIISAppPoolRecycleLoggingFix.log"
            $table=$null
          }
          finally
          {
            $connection.close()

          }
     }

try
{

	$HostNames = GetServerList
	 if ($HostNames -ne $null)
     {     
	  	for ($i=0; $i -lt $HostNames.Length; $i++)
        {
                $s = $HostNames[$i].CSHostname;
				$PoolName = $HostNames[$i].poolname;
				if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet)
                {
					if (Test-WSMan -ComputerName $s)
                    {
						$Error.Clear();
	    				$res = Invoke-command -computer $s -ArgumentList $s,$PoolName -ScriptBlock $ToExecute -ErrorAction SilentlyContinue -credential $DBCredential -SessionOption (New-PSSessionOption -OperationTimeout 90000)
						if ($Error.Count -gt 0)
						{
							  $PoolLoggingObj = New-Object -TypeName PsObject -Property(@{
						            'Server'  = $s;
									'PoolName'  = $PoolName;
									'Status'  = "ERROR";
									'Exception'= $Error[0];
						       })
							$Insert = InsertServerIISAppPoolRecycleLoggingFix -CSHostname $PoolLoggingObj.Server -PoolName $PoolLoggingObj.PoolName -Status $PoolLoggingObj.Status -Exception $PoolLoggingObj.Exception
						}
						else
						{
							$Insert = InsertServerIISAppPoolRecycleLoggingFix -CSHostname $res.Server -PoolName $res.PoolName -Status $res.Status -Exception $res.Exception
							
						}
					}else
					{
					$PoolLoggingObj = New-Object -TypeName PsObject -Property(@{
						            'Server'  = $s;
									'PoolName'  = $PoolName;
									'Status'  = "ERROR";
									'Exception'= "ERROR: The client cannot connect to the destination specified in the request";
						       })
						$Insert = InsertServerIISAppPoolRecycleLoggingFix -CSHostname $PoolLoggingObj.Server -PoolName $PoolLoggingObj.PoolName -Status $PoolLoggingObj.Status -Exception $PoolLoggingObj.Exception
					}
				}else
				{
					$PoolLoggingObj = New-Object -TypeName PsObject -Property(@{
						            'Server'  = $s;
									'PoolName'  = $PoolName;
									'Status'  = "ERROR";
									'Exception'= "ERROR: Server is not Reachable";
						       })
					
						$Insert = InsertServerIISAppPoolRecycleLoggingFix -CSHostname $PoolLoggingObj.Server -PoolName $PoolLoggingObj.PoolName -Status $PoolLoggingObj.Status -Exception $PoolLoggingObj.Exception
				}
		}
	 }else
     {
       WriteLog -msg "There is an error when trying to get Server List from database" -logname "ServerIISAppPoolRecycleLoggingFix.log"
     }
}catch
{
	$PoolLoggingObj = New-Object -TypeName PsObject -Property(@{
		'Server'  = $s;
		'PoolName'  = $PoolName;
		'Status'  = "ERROR";
		'Exception'= "ERROR: $_.Exception.Message";
	})
	$Insert = InsertServerIISAppPoolRecycleLoggingFix -CSHostname $PoolLoggingObj.Server -PoolName $PoolLoggingObj.PoolName -Status $PoolLoggingObj.Status -Exception $PoolLoggingObj.Exception			
	
}

#Sending E-mail
SendEmail -logname "ServerIISAppPoolRecycleLoggingFix.log"
WriteLog -msg "Process has Finished."	-logname "ServerIISAppPoolRecycleLoggingFix.log"
