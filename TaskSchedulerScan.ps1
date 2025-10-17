# Scan task schedulers using Powershell JOBs
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
      
	  $schtask = schtasks.exe /query /s $s  /V /FO CSV | ConvertFrom-Csv  
	  $TaskList = $schtask | where {$_."Run As User" -ne "NT AUTHORITY\LOCAL SERVICE" -and $_."Run As User" -ne "NT AUTHORITY\NETWORK SERVICE" -and $_."Run As User" -ne "NT AUTHORITY\SYSTEM" -and $_."Run As User" -ne "Run As User" -and $_."Run As User" -ne "BUILTIN\Users" -and $_."Run As User" -ne "\Everyone" -and $_."Run As User" -ne "NT AUTHORITY\INTERACTIVE" -and $_."Run As User" -ne "NT AUTHORITY\Authenticated Users" -and $_."Run As User" -ne "BUILTIN\Administrators"-and $_."Run As User" -ne "\" -and $_."Run As User" -notlike "*\CREATOR*" -and $_."Run As User" -notlike "*\ADM*" -and $_."TaskName" -notlike "\Optimize Start Menu Cache Files*" -and $_."TaskName" -notlike "\User_Feed_Synchronization*" -and $_."TaskName" -notlike "\WPD\SqmUpload*" -and $_."Run As User" -ne "AMERICAS\Domain Users"-and $_."Run As User" -ne "ASIA-PACIFIC\Domain Users" -and $_."Run As User" -ne "EUROPE\Domain Users"-and $_."Run As User" -ne "JAPAN\Domain Users" -and $_."Run As User" -ne "SYSTEM" -and $_."Run As User" -ne "Users" -and $_."Run As User" -ne "NETWORK SERVICE" -and $_."Run As User" -ne "LOCAL SERVICE" -and $_."Run As User" -ne "INTERACTIVE" -and $_."Run As User" -ne "Authenticated Users" -and $_."Run As User" -ne "Administrators"}

	  
	  if ($TaskList -eq $null)	
	  {
	    $TaskObj = New-Object -TypeName PsObject -Property(@{
                    'Server'  = $s;
                    'TaskName' =   "Not Found";
                    'Status'=  "Not Found";
                    'TaskState'=  "Not Found";
                    'User' ="Not Found";
                    'LastResult' =  "Not Found";
				          	'NextRunTime' = "Not Found";
					          'LastRunTime' = "Not Found";
                   
                     })
                     return $TaskObj
 
	  }else
	  {
        $TaskObj = New-Object -TypeName PsObject -Property(@{
                        'Server'  = $s;
                        'TaskName'= $TaskList;
                     })
                     return $TaskObj

	   
      }     

    }catch
    {
        $TaskObj = New-Object -TypeName PsObject -Property(@{
	                    'Server'  = $s;
	                    'TaskName' =   "ERROR";
	                    'Status'=  "ERROR";
	                    'TaskState'=  "ERROR";
	                    'User' = "ERROR";
	                    'LastResult' =  $_.Exception.Message;
						          'NextRunTime' =  "ERROR";
						          'LastRunTime' =  "ERROR";
                     })
                     return $TaskObj

    }



  }
}

   Function InsertTaskSchedulerData([String] $CSHostname, [String]$TaskName, [String]$TaskStatus, [String]$TaskState, [String]$TaskUser, [String]$TaskLastResult,[String]$NextRunTime,[String]$LastRunTime  )
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Task: $TaskName, Status: $TaskStatus, State: $TaskState, TaskUser: $TaskUser, Last Result: $TaskLastResult, NextRunTime: $NextRunTime, LastRunTime: $LastRunTime" -logname "TaskSchedulerScan.log"

          $Query = "INSERT INTO TaskScheduler (CSHostname, TaskName,TaskStatus,TaskState,TaskUser,TaskLastResult,NextRunTime,LastRunTime) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')" -f $CSHostname,$TaskName,$TaskStatus,$TaskState,$TaskUser,$TaskLastResult,$NextRunTime,$LastRunTime
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "TaskSchedulerScan.log"
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
           $delete = DeleteAllDataTable -tablename "TaskScheduler" -databasename "ServerAudit" -logname "TaskSchedulerScan.log"
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
                        $TaskObj = New-Object -TypeName PsObject -Property(@{
	                        'Server'  = $s;
							'TaskName' =   "ERROR";
		                    'Status'=  "ERROR";
		                    'TaskState'=  "ERROR";
		                    'User' = "ERROR";
		                    'LastResult' = "ERROR: The client cannot connect to the destination specified in the request";
							'NextRunTime' =  "ERROR";
							'LastRunTime' =  "ERROR";
	                   })
					   $insert = InsertTaskSchedulerData -CSHostname $TaskObj.Server -TaskName $TaskObj.TaskName -TaskStatus $TaskObj.Status -TaskState $TaskObj.TaskState -TaskUser $TaskObj.User -TaskLastResult $TaskObj.LastResult -NextRunTime $TaskObj.NextRunTime -LastRunTime $TaskObj.LastRunTime
                     }
                }else
                {
                   $TaskObj = New-Object -TypeName PsObject -Property(@{
	                        'Server'  = $s;
							'TaskName' =   "ERROR";
		                    'Status'=  "ERROR";
		                    'TaskState'=  "ERROR";
		                    'User' = "ERROR";
		                    'LastResult' = "ERROR: Server is not Reachable";
							'NextRunTime' =  "ERROR";
							'LastRunTime' =  "ERROR";
	                   })
					   $insert = InsertTaskSchedulerData -CSHostname $TaskObj.Server -TaskName $TaskObj.TaskName -TaskStatus $TaskObj.Status -TaskState $TaskObj.TaskState -TaskUser $TaskObj.User -TaskLastResult $TaskObj.LastResult -NextRunTime $TaskObj.NextRunTime -LastRunTime $TaskObj.LastRunTime
                }
            }
            
        }
        else
        {
            WriteLog -msg "List of Servers is empty" -logname "TaskSchedulerScan.log"
        }
    }catch
    {
        $TaskObj = New-Object -TypeName PsObject -Property(@{
	                      'Server'  = $s;
							          'TaskName' =   "ERROR";
		                    'Status'=  "ERROR";
		                    'TaskState'=  "ERROR";
		                    'User' = "ERROR";
		                    'LastResult' = $_.Exception.Message;
							          'NextRunTime' =  "ERROR";
							          'LastRunTime' =  "ERROR";
	                   })
					   $insert = InsertTaskSchedulerData -CSHostname $TaskObj.Server -TaskName $TaskObj.TaskName -TaskStatus $TaskObj.Status -TaskState $TaskObj.TaskState -TaskUser $TaskObj.User -TaskLastResult $TaskObj.LastResult -NextRunTime $TaskObj.NextRunTime -LastRunTime $TaskObj.LastRunTime
    }
    Get-Job | Wait-Job -Timeout 180
    $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName 
    
    foreach ($r in $results){

        $TaskList = $r.TaskName
    
        if ($TaskList -eq "ERROR" -or $TaskList -eq "Not Found")
        {
           $insert = InsertTaskSchedulerData -CSHostname $r.Server -TaskName $r.TaskName -TaskStatus $r.Status -TaskState $r.TaskState -TaskUser $r.User -TaskLastResult $r.LastResult -NextRunTime $r.NextRunTime -LastRunTime $r.LastRunTime 
        }else
        {
            foreach ($task in $TaskList)
            {
				$taskHost=$task.HostName
				if ($taskHost -notlike '*ERROR*')
				{
                 $insert = InsertTaskSchedulerData -CSHostname $r.Server -TaskName $task.TaskName -TaskStatus $task.Status -TaskState $task.'Scheduled Task State' -TaskUser $task.'Run As User' -TaskLastResult $task.'Last Result' -NextRunTime $task.'Next Run Time' -LastRunTime $task.'Last Run Time'
				}
				else{
				    WriteLog -msg "Error: There is a problem with a task scheduler: Error Details:$taskHost" -logname "TaskSchedulerScan.log"
				}
            }
        }
    }
     
        #Sending E-mail
        SendEmail -logname "TaskSchedulerScan.log"
        WriteLog -msg "Process has Finished." -logname "TaskSchedulerScan.log"
}