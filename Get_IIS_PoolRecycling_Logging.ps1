<#
.DESCRIPTION
    1)	You can use this script to get details whether pools are logging when a recycle happens in event logs


.NOTES
    File Name      : Get_IIS_PoolReclycing_Logging.ps1
    Author         : Rodrigo Scorsatto (Rodrigo_Scorsatto@dell.com)
    Prerequisite   : PowerShell V2 or upper
    Version        : 1.0 - 08/08/2019

#>


$ToExecute =
{
     param([string] $server)  
    
   
   try{
    
    Import-Module WebAdministration
    #Add-PSSnapin WebAdministration
    
    #get a list of all Active app Pools
    $Pool = Get-ChildItem IIS:\AppPools #| ? {$_.state -eq "Started"}

    $result = @()
    foreach ( $p in $Pool )
    {
        $app_pool = $p.Name
        $pool_status = $p.State
        $IIS = Get-ItemProperty "IIS:\AppPools\$app_pool"
		
		$userName = $IIS.processModel.userName
		
		if ($userName -eq "")
		{
			$userName="ApplicationPoolIdentity"
		}
		
        $flags = $IIS.Recycling.Attributes["logEventOnRecycle"].Value

        $RegularTimeInterval = [bool]($flags -band 1)     # Regular Time Interval
        $requests = [bool]($flags -band 2)     # Request Limit Exceeded
        $schedule = [bool]($flags -band 4)     # Specific Time
        $memory         = [bool]($flags -band 8)     # Virtual Memory Limit Exceeded
        $isapiUnhealthy = [bool]($flags -band 16)    # Isapi Reported Unhealthy
        $onDemand      = [bool]($flags -band 32)    # Manual Recycle
        $configChange   = [bool]($flags -band 64)    # Application Pool Configuration Changed
        $privateMemory  = [bool]($flags -band 128)   # Private Memory Limit Exceeded
        
		#Dot Net Core Info
		if ($IIS.managedRuntimeVersion -eq "")
		{
			$managedRuntimeVersion = "undefined"
		}else
		{
			$managedRuntimeVersion = $IIS.managedRuntimeVersion
		}
		
		#Recycle Options Info
		
		$recycleinfo = Get-itemproperty -Path IIS:\AppPools\$app_pool -Name recycling.periodicRestart
		$RecycleRegularTime = $recycleinfo.time.TotalMinutes
		$RecycleNumberRequets = $recycleinfo.requests
		$RecycleVirtualMemoryUsage = $recycleinfo.memory
	 	$RecyclePrivateMemoryUsage = $recycleinfo.privatememory
		
		#Specific Times
		$ScheduledTimesWindow = $recycleinfo.schedule.Collection
		$TimesResults = $null
		foreach ($times in $ScheduledTimesWindow)
	 	{
			$time=$times.value.Tostring()
			$time = $time+" | "
			$TimesResults+=$time;
		}

		
        $RecycleObj = New-Object -TypeName PsObject -Property(@{
            'Server'  = $server;
            'PoolName'= $app_pool;
            'Status' = $pool_status;
            'Regular Time Interval'    = $RegularTimeInterval;
            'Request Limit Exceeded'= $requests;
            'Specific Time'  = $schedule;
            'Virtual Memory Limit Exceeded' = $memory;
            'Isapi Reported Unhealthy' = $isapiUnhealthy;
            'Manual Recycle'       = $onDemand;
            'Application Pool Configuration Changed'   = $configChange;
            'Private Memory Limit Exceeded'  = $privateMemory;
			'ManagedRuntimeVersion'= $managedRuntimeVersion;
			'RecycleRegularTime'  = $RecycleRegularTime;
			'$RecycleNumberRequets'  = $RecycleNumberRequets;
			'RecycleVirtualMemoryUsage'  = $RecycleVirtualMemoryUsage;
			'RecyclePrivateMemoryUsage'  = $RecyclePrivateMemoryUsage;
			'RecycleSpecificTimes'  = $TimesResults;
			'Identity' = $userName;
        })
        $result += $RecycleObj
    }
     return $result
    }catch
    {
        $exception =  $_.Exception.Message;
         $RecycleObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $server;
                        'PoolName'= "ERROR";
                        'Status' = "ERROR";
                        'Regular Time Interval'    = "ERROR";
                        'Request Limit Exceeded'= "ERROR";
                        'Specific Time'  = "ERROR";
                        'Virtual Memory Limit Exceeded' = "ERROR";
                        'Isapi Reported Unhealthy' = "ERROR";
                        'Manual Recycle'       = "ERROR";
                        'Application Pool Configuration Changed'   = "ERROR";
                        'Private Memory Limit Exceeded'  = [String]$exception;
						'ManagedRuntimeVersion'= "ERROR";
						'RecycleRegularTime'  = "ERROR";
						'$RecycleNumberRequets'  = "ERROR";
						'RecycleVirtualMemoryUsage'  = "ERROR";
						'RecyclePrivateMemoryUsage'  = "ERROR";
						'RecycleSpecificTimes'  = "ERROR";
						'Identity' = "ERROR";
           })
           
        return $RecycleObj
    
    }
}


     Function InsertServerIISAppPoolRecycleLoggingData([String] $CSHostname, [String]$PoolName, [String] $Status, [String]$RegularTimeInterval, [String]$RequestLimitExceeded, [String]$SpecificTime, [String]$VirtualMemoryLimitExceeded, [String]$IsapiReportedUnhealthy, [String] $ManualRecycle, [String] $ApplicationPoolConfigurationChanged,[String] $PrivateMemoryLimitExceeded,[String]$ManagedRuntimeVersion,[String]$RecycleRegularTime,[String]$RecycleNumberRequets,[String]$RecycleVirtualMemoryUsage,[String]$RecyclePrivateMemoryUsage,[String]$RecycleSpecificTimes, [String]$Identity)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Pool: $PoolName, Status: $Status, RegularTimeInterval: $RegularTimeInterval,RequestLimitExceeded: $RequestLimitExceeded, SpecificTime: $SpecificTime, VirtualMemoryLimitExceeded: $VirtualMemoryLimitExceeded, IsapiReportedUnhealthy: $IsapiReportedUnhealthy, ManualRecycle: $ManualRecycle, ApplicationPoolConfigurationChanged: $ApplicationPoolConfigurationChanged, PrivateMemoryLimitExceeded: $PrivateMemoryLimitExceeded, Identity: $Identity" -logname "GetIIS_PoolRecycleLogging.log"

          $Query = "INSERT INTO ServerIISAppPoolRecycleLogging (CSHostname, PoolName,Status,RegularTimeInterval,RequestLimitExceeded,SpecificTime,VirtualMemoryLimitExceeded,IsapiReportedUnhealthy, ManualRecycle, ApplicationPoolConfigurationChanged, PrivateMemoryLimitExceeded,ManagedRuntimeVersion, RecycleRegularTime, RecycleNumberRequets,RecycleVirtualMemoryUsage,RecyclePrivateMemoryUsage,RecycleSpecificTimes, [Identity]) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}')" -f $CSHostname,$PoolName, $Status, $RegularTimeInterval, $RequestLimitExceeded, $SpecificTime, $VirtualMemoryLimitExceeded, $IsapiReportedUnhealthy, $ManualRecycle, $ApplicationPoolConfigurationChanged, $PrivateMemoryLimitExceeded,$ManagedRuntimeVersion, $RecycleRegularTime, $RecycleNumberRequets,$RecycleVirtualMemoryUsage,$RecyclePrivateMemoryUsage,$RecycleSpecificTimes, $Identity
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
            $exception =  $_.Exception.Message;
            WriteLog -msg "There is an error when trying to insert data. Error: $exception" -logname "GetIIS_PoolRecycleLogging.log"
            $result=$null
          }
          finally
          {
            $conn.close()

          }
     }


try{
    
        $HostNames = GeServerList_IIS_Only
        
         if ($HostNames -ne $null)
         { 
           $delete = DeleteAllDataTable -tablename "ServerIISAppPoolRecycleLogging" -databasename "ServerAudit" -logname "GetIIS_PoolRecycleLogging.log"
           for ($i=0; $i -lt $HostNames.Length; $i++)
           {
                $s = $HostNames[$i].CSHostname;
                
                if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet)
                {
                    if (Test-WSMan -ComputerName $s)
                    {
						$Error.Clear();
                        $res = Invoke-command -computer $s -ScriptBlock $ToExecute -ArgumentList $s -credential $ServerCredential -SessionOption (New-PSSessionOption -OperationTimeout 90000)
                        if ($Error.Count -gt 0)
						{
						 
						    $RecycleObj = New-Object -TypeName PsObject -Property(@{
		                        'Server'  = $s;
		                        'PoolName'= "ERROR";
                                'Status' = "ERROR";
		                        'Regular Time Interval'    = "ERROR";
		                        'Request Limit Exceeded'= "ERROR";
		                        'Specific Time'  = "ERROR";
		                        'Virtual Memory Limit Exceeded' = "ERROR";
		                        'Isapi Reported Unhealthy' = "ERROR";
		                        'Manual Recycle'       = "ERROR";
		                        'Application Pool Configuration Changed'   = "ERROR";
		                        'Private Memory Limit Exceeded'  = [String]$Error[0];
								'ManagedRuntimeVersion'= "ERROR";
								'RecycleRegularTime'  = "ERROR";
								'$RecycleNumberRequets'  = "ERROR";
								'RecycleVirtualMemoryUsage'  = "ERROR";
								'RecyclePrivateMemoryUsage'  = "ERROR";
								'RecycleSpecificTimes'  = "ERROR";
								'Identity' ="ERROR";
		                     })
                     $insert = InsertServerIISAppPoolRecycleLoggingData -CSHostname $RecycleObj.Server -PoolName $RecycleObj.PoolName -Status $RecycleObj.Status -RegularTimeInterval $RecycleObj.'Regular Time Interval' -RequestLimitExceeded $RecycleObj.'Request Limit Exceeded' -SpecificTime $RecycleObj.'Specific Time' -VirtualMemoryLimitExceeded $RecycleObj.'Virtual Memory Limit Exceeded' -IsapiReportedUnhealthy $RecycleObj.'Isapi Reported Unhealthy' -ManualRecycle $RecycleObj.'Manual Recycle' -ApplicationPoolConfigurationChanged $RecycleObj.'Application Pool Configuration Changed' -PrivateMemoryLimitExceeded $RecycleObj.'Private Memory Limit Exceeded' -Identity $RecycleObj.Identity
      			               	
						}
						else
						{
							foreach ($result in $res)
	                        {
	                            #$res | Select-Object Server, PoolName, 'Regular Time Interval', 'Request Limit Exceeded', 'Specific Time', 'Virtual Memory Limit Exceeded', 'Isapi Reported Unhealthy', 'Manual Recycle', 'Application Pool Configuration Changed', 'Private Memory Limit Exceeded' | Export-Csv $path  -Append -NoTypeInformation
	                            $insert = InsertServerIISAppPoolRecycleLoggingData -CSHostname $result.Server -PoolName $result.PoolName -Status $result.Status -RegularTimeInterval $result.'Regular Time Interval' -RequestLimitExceeded $result.'Request Limit Exceeded' -SpecificTime $result.'Specific Time' -VirtualMemoryLimitExceeded $result.'Virtual Memory Limit Exceeded' -IsapiReportedUnhealthy $result.'Isapi Reported Unhealthy' -ManualRecycle $result.'Manual Recycle' -ApplicationPoolConfigurationChanged $result.'Application Pool Configuration Changed' -PrivateMemoryLimitExceeded $result.'Private Memory Limit Exceeded' -ManagedRuntimeVersion $result.ManagedRuntimeVersion -RecycleRegularTime $result.RecycleRegularTime -RecycleNumberRequets $result.RecycleNumberRequets -RecycleVirtualMemoryUsage $result.RecycleVirtualMemoryUsage -RecyclePrivateMemoryUsage $result.RecyclePrivateMemoryUsage -RecycleSpecificTimes $result.RecycleSpecificTimes -Identity $result.Identity
	                        }
						}
                     }else
                     {
                        $RecycleObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'PoolName'= "ERROR";
                        'Status' = "ERROR";
                        'Regular Time Interval'    = "ERROR";
                        'Request Limit Exceeded'= "ERROR";
                        'Specific Time'  = "ERROR";
                        'Virtual Memory Limit Exceeded' = "ERROR";
                        'Isapi Reported Unhealthy' = "ERROR";
                        'Manual Recycle'       = "ERROR";
                        'Application Pool Configuration Changed'   = "ERROR";
                        'Private Memory Limit Exceeded'  = "ERROR: The client cannot connect to the destination specified in the request";
						'ManagedRuntimeVersion'= "ERROR";
						'RecycleRegularTime'  = "ERROR";
						'$RecycleNumberRequets'  = "ERROR";
						'RecycleVirtualMemoryUsage'  = "ERROR";
						'RecyclePrivateMemoryUsage'  = "ERROR";
						'RecycleSpecificTimes'  = "ERROR";
						'Identity' ="ERROR";
                     })
                     $insert = InsertServerIISAppPoolRecycleLoggingData -CSHostname $RecycleObj.Server -PoolName $RecycleObj.PoolName -Status $RecycleObj.Status -RegularTimeInterval $RecycleObj.'Regular Time Interval' -RequestLimitExceeded $RecycleObj.'Request Limit Exceeded' -SpecificTime $RecycleObj.'Specific Time' -VirtualMemoryLimitExceeded $RecycleObj.'Virtual Memory Limit Exceeded' -IsapiReportedUnhealthy $RecycleObj.'Isapi Reported Unhealthy' -ManualRecycle $RecycleObj.'Manual Recycle' -ApplicationPoolConfigurationChanged $RecycleObj.'Application Pool Configuration Changed' -PrivateMemoryLimitExceeded $RecycleObj.'Private Memory Limit Exceeded' -Identity $RecycleObj.Identity
                     
                     }
                }else
                {
                    $exception =  $_.Exception.Message;
                    $RecycleObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'PoolName'= "ERROR";
                        'Status' = "ERROR";
                        'Regular Time Interval'    = "ERROR";
                        'Request Limit Exceeded'= "ERROR";
                        'Specific Time'  = "ERROR";
                        'Virtual Memory Limit Exceeded' = "ERROR";
                        'Isapi Reported Unhealthy' = "ERROR";
                        'Manual Recycle'       = "ERROR";
                        'Application Pool Configuration Changed'   = "ERROR";
                        'Private Memory Limit Exceeded'  = [String]$exception;
						'ManagedRuntimeVersion'= "ERROR";
						'RecycleRegularTime'  = "ERROR";
						'$RecycleNumberRequets'  = "ERROR";
						'RecycleVirtualMemoryUsage'  = "ERROR";
						'RecyclePrivateMemoryUsage'  = "ERROR";
						'RecycleSpecificTimes'  = "ERROR";
						'Identity' ="ERROR";
                     })
                     $insert = InsertServerIISAppPoolRecycleLoggingData -CSHostname $RecycleObj.Server -PoolName $RecycleObj.PoolName -Status $RecycleObj.Status -RegularTimeInterval $RecycleObj.'Regular Time Interval' -RequestLimitExceeded $RecycleObj.'Request Limit Exceeded' -SpecificTime $RecycleObj.'Specific Time' -VirtualMemoryLimitExceeded $RecycleObj.'Virtual Memory Limit Exceeded' -IsapiReportedUnhealthy $RecycleObj.'Isapi Reported Unhealthy' -ManualRecycle $RecycleObj.'Manual Recycle' -ApplicationPoolConfigurationChanged $RecycleObj.'Application Pool Configuration Changed' -PrivateMemoryLimitExceeded $RecycleObj.'Private Memory Limit Exceeded' -Identity $RecycleObj.Identity
                }
            }
            
        }
        else
        {
            WriteLog -msg "List of Servers is empty" -logname "GetIIS_PoolRecycleLogging.log"
        }
    }catch
    {
      $exception =  $_.Exception.Message;
      $RecycleObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'PoolName'= "ERROR";
                        'Status' = "ERROR";
                        'Regular Time Interval'    = "ERROR";
                        'Request Limit Exceeded'= "ERROR";
                        'Specific Time'  = "ERROR";
                        'Virtual Memory Limit Exceeded' = "ERROR";
                        'Isapi Reported Unhealthy' = "ERROR";
                        'Manual Recycle'       = "ERROR";
                        'Application Pool Configuration Changed'   = "ERROR";
                        'Private Memory Limit Exceeded'  = [String]$exception;
						'ManagedRuntimeVersion'= "ERROR";
						'RecycleRegularTime'  = "ERROR";
						'$RecycleNumberRequets'  = "ERROR";
						'RecycleVirtualMemoryUsage'  = "ERROR";
						'RecyclePrivateMemoryUsage'  = "ERROR";
						'RecycleSpecificTimes'  = "ERROR";
						'Identity' ="ERROR";
                     })
                     InsertServerIISAppPoolRecycleLoggingData -CSHostname $RecycleObj.Server -PoolName $RecycleObj.PoolName -Status $RecycleObj.Status -RegularTimeInterval $RecycleObj.'Regular Time Interval' -RequestLimitExceeded $RecycleObj.'Request Limit Exceeded' -SpecificTime $RecycleObj.'Specific Time' -VirtualMemoryLimitExceeded $RecycleObj.'Virtual Memory Limit Exceeded' -IsapiReportedUnhealthy $RecycleObj.'Isapi Reported Unhealthy' -ManualRecycle $RecycleObj.'Manual Recycle' -ApplicationPoolConfigurationChanged $RecycleObj.'Application Pool Configuration Changed' -PrivateMemoryLimitExceeded $RecycleObj.'Private Memory Limit Exceeded' -Identity $RecycleObj.Identity
    }

     
        #Sending E-mail
        SendEmail -logname "GetIIS_PoolRecycleLogging.log"
        WriteLog -msg "Process has Finished." -logname "GetIIS_PoolRecycleLogging.log"