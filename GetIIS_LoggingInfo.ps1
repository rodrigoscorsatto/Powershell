<#
.DESCRIPTION
    1)	You can use this script to get IIS Logging Path and Failed Trace Requests info.


.NOTES
    File Name      : GetIIS_LoggingInfo.ps1
    Author         : Rodrigo Scorsatto 
    Prerequisite   : PowerShell V2 or upper
    Version        : 1.0 - 08/14/2019

#>



$sbCommand = 
        {
            param ([Parameter(Mandatory=$true,ValueFromPipeline=$true)][string] $Server);
        
            $startDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
            function Get-IIS-Version {
	                    [Cmdletbinding()]
                        param([String]$Server)
                        try
                        {
                            $reg = Get-WmiObject -List StdRegProv -Namespace root\default -ErrorAction Stop
                            if ($reg) {
			                    $HKLM = 2147483650
			                    $regVer = $reg.GetStringValue($HKLM,"SOFTWARE\Microsoft\InetStp","VersionString").sValue
			                    if ($regVer){
		    	                    return $regVer.replace('Version','').trim()
			                    }
			                    else{ return "No IIS identified" }
		                    }
		                    else{return "ERROR - Could not collect $Server information from registry"}
	                    } #end of try
	                    catch {
		                    $Message = $_.Exception.Message
    	                    $FailedItem = $_.Exception.ItemName
		                    $return = "Error " + (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name + " server $Server - $Message  - $FailedItem"
		                    $return
	                    }
                    }
                    $versionIIS = $(Get-IIS-Version $env:COMPUTERNAME).trim()
                    if ($versionIIS -match "7.*" -or $versionIIS -match "8.*" -or $versionIIS -match "9.*" -or $versionIIS -match "10.*" -or $versionIIS -match "11.*")
                    {
                        try
                        {
                            #read all IIS configs
                            $xml = [xml]$(Get-Content -Path "C:\Windows\System32\inetsrv\config\applicationHost.config")
                            if (![string]::IsNullOrEmpty($xml)){
                                try{$xmlSites = $xml.GetElementsByTagName("site")}
                                catch{}
                            }

                            $PathAppHostConfig = "C:\Windows\System32\inetsrv\config\schema\IIS_Schema.xml" -replace [regex]::escape(':'),'$'
                            [xml]$IIS_Schema = [xml]$(Get-Content -Path "C:\Windows\System32\inetsrv\config\schema\IIS_Schema.xml")
                            $Trace = @{}
                            if (![string]::IsNullOrEmpty($IIS_Schema)){
                                try{ 
                                    $($($($IIS_Schema.GetElementsByTagName("sectionSchema") | Where-Object{$_.name -ieq "system.applicationHost/sites"}).GetElementsByTagName("collection") | Where-Object {$_.addElement -ieq "site" -and $_.defaultElement -ieq "siteDefaults"}).GetElementsByTagName("element") | Where-Object {$_.name -ieq "traceFailedRequestsLogging"}).attribute | select name,defaultValue | foreach{ $Trace[$_.name] = $_.defaultValue} 
                                    }
                                catch{}
                            }

                            try
                            {
                                [array]$allSites = @()
                                [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Web.Administration") | Out-Null
                                $serverManager = new-object Microsoft.Web.Administration.ServerManager
                                [array]$allSites = $serverManager.Sites #| select id,state
                                [array]$allApplicationPools = $serverManager.ApplicationPools # | select id,state
                                $serverManager.Dispose()
                            }
                            catch{}

                            $Sites = $xml.GetElementsByTagName("sites")
			                try
			                {
                                [string]$default_LogPath = $Sites.GetElementsByTagName("siteDefaults").logFile.directory
                                [string]$default_appPool = $Sites.GetElementsByTagName("applicationDefaults").applicationPool

                                [string]$LogPathEnabled_Default = $xml.GetElementsByTagName("siteDefaults").logFile.enabled
                                [string]$LogPathXML_Default = $xml.GetElementsByTagName("siteDefaults").logFile.directory
                                [string]$TraceFailedRequestsLogPath_Default = $xml.GetElementsByTagName("siteDefaults").traceFailedRequestsLogging.directory -ireplace '%SystemDrive%','C:'
			                }
			                catch {
				            [string]$default_LogPath = $sites | select -ExpandProperty siteDefaults | select -ExpandProperty logFile | select -ExpandProperty directory
				            [string]$default_appPool = $sites | select -ExpandProperty applicationDefaults | select -ExpandProperty applicationPool

				            try{
					            [string]$LogPathEnabled_Default = $($xml.GetElementsByTagName("siteDefaults") | select -ExpandProperty logFile | select enabled).enabled
					            if ([string]::IsNullorEmpty($LogPathEnabled_Default)){[string]$LogPathEnabled_Default = 'unidentified'}
					
					            [string]$LogPathXML_Default = $($xml.GetElementsByTagName("siteDefaults") | select -ExpandProperty logFile | select directory).directory
					            if ([string]::IsNullorEmpty($LogPathXML_Default)){[string]$LogPathXML_Default = 'unidentified'}

				            }
				            catch{[string]$LogPathEnabled_Default = 'unidentified'}

				            try{[string]$TraceFailedRequestsLogPath_Default = $xml.GetElementsByTagName("siteDefaults") | select -ExpandProperty traceFailedRequestsLogging | select -ExpandProperty directory}
				            catch{[string]$TraceFailedRequestsLogPath_Default = 'unidentified'}

                            }
                            foreach ($w in $allSites)
                            {
                                if ([array]$allSites.count -eq 0)
                                {
                                    [string]$siteState = 'Unable to identify state'
                                }
                                else
                                {
                                    [string]$siteState = $($allSites | Where{$_.Id -eq $w.id}).State
                                    if ([string]::IsNullOrEmpty($siteState))
                                    {
                                        $siteState = 'Unable to identify state'
                                    }
                                }
                                $site = $($xml.GetElementsByTagName("site") | Where-Object {$_.name -ieq $w.name} | Sort-Object -Unique)
								
								$PhysicalPath = $null
                                $PhysicalPath = $w.Applications.VirtualDirectories.PhysicalPath
								if ($PhysicalPath -ne $null -or $PhysicalPath -eq '')
								{
									$PhysicalPath = $PhysicalPath -ireplace ('%SystemDrive%','C:')
									foreach ($path in $PhysicalPath)
									{
										$PhysicalPathList+= $path+";"
									}
									
									$Paths = $PhysicalPathList -split(";")

									$PathList=@()
									foreach ($path in $Paths)
									{
									  $selection = ($path  | Select-String "\\" -AllMatches)
									  $path = $path.tolower()   
									  
									  #special condition for paths like \\server\paths
									  $RemotePath = ($path  | Select-String "\\\\" -AllMatches)
									  $RemoteCount = $RemotePath.Matches.Count
									  $Counter = $selection.Matches.Count
									  $Counter = $Counter - $RemoteCount
									 
									  if ($Counter -gt 2)
									  {
									    if ( ($Counter % 2) -eq 0 )
									    { 
											$Nrem =$Counter/2;
									    }else
										{ 
									      	$Nrem =[math]::truncate($Counter/2)+1;
									    }	
										
									    while ($Counter -ne $Nrem)
									    {
											$path = $path | Split-Path 
											$Nrem++	
									   }
									  }
									  $path=$path+";"
									  $PathList+=$path
									}
									$PathList = $PathList | Select-Object -Unique #-SkipLast 1
								}else
								{
									$PathList = "Not Found - Older IIS version"
								}

                                [string]$LogPath = $( if( [string]::IsNullOrEmpty($($w.logFile.directory -ireplace '%SystemDrive%','C:'))) {$($LogPathXML_Default -ireplace '%SystemDrive%','C:')} else {$($w.logFile.directory -ireplace '%SystemDrive%','C:')})
                                [string]$LogEnabled = $(if ([string]::IsNullOrEmpty($site.logFile.enabled) ){ $($xml.GetElementsByTagName("siteDefaults").logFile.enabled) } else { $site.logFile.enabled })
                                [string]$maxLogFileSizeKB = $(if ([string]::IsNullOrEmpty($site.maxLogFileSizeKB)){ $Trace.maxLogFileSizeKB } else { $site.maxLogFileSizeKB })
                                [string]$TraceFailedRequests = $(if ([string]::IsNullOrEmpty($site.traceFailedRequestsLogging.enabled)){ $Trace.enabled } else { $site.traceFailedRequestsLogging.enabled })
                                [string]$TraceFailedRequestsLogPath = $(if ([string]::IsNullOrEmpty($($site.TraceFailedRequestsLogging.Directory -ireplace '%SystemDrive%','C:'))){if([string]::IsNullOrEmpty([string]$TraceFailedRequestsLogPath_Default)){$($Trace.directory -ireplace '%SystemDrive%','C:')}else{[string]$TraceFailedRequestsLogPath_Default}} else { $($site.TraceFailedRequestsLogging.Directory -ireplace '%SystemDrive%','C:') })
                                [string]$MaxLogFiles = $(if ([string]::IsNullOrEmpty($site.TraceFailedRequestsLogging.maxLogFiles)){ $($Trace.maxLogFiles) } else { $site.TraceFailedRequestsLogging.maxLogFiles })
                                [string]$Binding = $($([array]$($w.bindings | where {$_.protocol -ilike '*http'}| Sort-Object @{e={ [int]$($_.BindingInformation -replace '[^0-9]')}} -unique | select @{Name='Binding';e={'http/' + $($_.BindingInformation -replace '[^0-9]')}} | select -ExpandProperty Binding) + [array]$($w.bindings | where {$_.protocol -ilike '*https'}| Sort-Object @{e={ [int]$($_.BindingInformation -replace '[^0-9]')}} -unique | select @{Name='Binding';e={'https/' + $($_.BindingInformation -replace '[^0-9]')}} | select -ExpandProperty Binding)) -join ',').Split(',',[System.StringSplitOptions]::RemoveEmptyEntries) -join ','
                                [string]$appPool = $(
                                                        if(![string]::IsNullOrEmpty($($($w | ForEach-Object {$_.Applications} | ForEach-Object {$_.ApplicationPoolName} | Sort-Object -Unique) -join [environment]::NewLine)))
                                                        {
                                                            $($($w | ForEach-Object {$_.Applications} | ForEach-Object {$_.ApplicationPoolName} | Sort-Object -Unique) -join [environment]::NewLine)
                                                        }
                                                        elseif ([string]::IsNullOrEmpty($($site.application.applicationpool -join [environment]::NewLine)))
                                                        {
                                                            if ([string]::IsNullOrEmpty($default_appPool)){"DefaultAppPool"} else {$default_appPool}
                                                        }
                                                        else
                                                        {
                                                            $($($site.application.applicationpool | Sort-Object -Unique) -join [environment]::NewLine)
                                                        }
                                                    )

                                [array]$return += [pscustomobject]@{
                                                            ServerName = $Server
							                                Connect = 1
                                                            AppPool = $appPool
                                                            SiteName = if (![string]::IsNullOrEmpty($site.name)){$site.name}elseif(![string]::IsNullOrEmpty($w.name)){$w.Name}else{'unidentified'}
                                                            State = $siteState
                                                            LogEnabled = $($(if ([string]::IsNullOrEmpty([string]![bool]$w.DontLog)){$LogEnabled}else{[string]![bool]$w.DontLog}))
                                                            LogPath = [string]$($LogPath + "\W3SVC"+ $w.id)
                                                            Binding = $Binding
                                                            MaxLogFiles = $MaxLogFiles
                                                            MaxLogFileSizeKB = $maxLogFileSizeKB
                                                            TraceFailedRequests = $($TraceFailedRequests)
                                                            TraceFailedRequestsLogPath = $TraceFailedRequestsLogPath
															PhysicalPath = $PathList
                                                            StartDate = ''
                                                            FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
			                                                }
                            }

                            if ($return.count -gt 0)
                            {
                                return $return
                            }
                            else
                            {
                        
                                [array]$return += [pscustomobject]@{
                                                        ServerName = $Server
						                                Connect = 1
                                                        SiteName = 'N/A - No site detected on IIS'
                                                        StartDate = ''
                                                        FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
		                                                }
                                return $return
                            }


                        }
                        catch
                        {
                            [array]$return += [pscustomobject]@{
                                                        ServerName = $Server
						                                Connect = 1
                                                        StartDate = ''
                                                        FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
                                                        ErrorMessage = "FAILED - Unable to detect IIS information: " + $_.Exception
		                                                }
                            return $return
                        }
                    }
                    else
                    {
				        $return += [pscustomobject]@{
							        ServerName = $Server
						            Connect = 1
							        SiteName = "N/A - No IIS 7 (or higher) detected"
                                    StartDate = $startDate
                                    FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
							        }
				            return $return
				    }

    }
                  

     Function InsertServerIISLoggingPath([String] $CSHostname, [String]$connect, [String]$AppPool, [String]$SiteName, [String]$State, [String]$LogEnabled, [String]$LogPath, [String]$Binding, [String]$MaxLogFiles, [String]$MaxLogFileSizeKB, [String]$TraceFailedRequests, [String]$TraceFailedRequestsLogPath, [String] $PhysicalPath)
     {
       try{
		  
          [String] $StringPath = $PhysicalPath
		  $StringPath = $StringPath.Replace("; ",";") #remove extra spaces
		  $StringPath = $StringPath.TrimEnd(";") #remove extra ; at end
		  
          WriteLog -msg "Inserting Server: $CSHostname with data: Connect: $connect, SiteName: $SiteName, LogEnabled: $LogEnabled, LogPath: $LogPath, TraceFailedRequests: $TraceFailedRequests, PhysicalPath: $StringPath" -logname "ServerIISLoggingPath.log"
          $Query = "INSERT INTO ServerIISLoggingPath (CSHostname, [connect], AppPool, SiteName, [State], LogEnabled, LogPath, [Binding], MaxLogFiles, MaxLogFileSizeKB, TraceFailedRequests, TraceFailedRequestsLogPath, PhysicalPath) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}')" -f $CSHostname,$connect,$AppPool,$SiteName,$State,$LogEnabled,$LogPath,$Binding,$MaxLogFiles,$MaxLogFileSizeKB, $TraceFailedRequests, $TraceFailedRequestsLogPath,$StringPath
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
            WriteLog -msg "There is an error when trying to insert data. Error: $exception" -logname "ServerIISLoggingPath.log"
            $result=$null
          }
          finally
          {
            $conn.close()

          }
     }

     


try
{
   
  
   $HostNames = GeServerList_IIS_Only

  if ($HostNames -ne $null)
  { 
           $delete = DeleteAllDataTable -tablename "ServerIISLoggingPath" -databasename "ServerAudit" -logname "ServerIISLoggingPath.log"
           for ($i=0; $i -lt $HostNames.Length; $i++)
           {
               
                $s = $HostNames[$i].CSHostname;   
                WriteLog -msg "Processing Server: $s" -logname "ServerIISLoggingPath.log"
              if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet)
              {
             try
             {

                if (Test-WSMan -ComputerName $s)
                {
                     $return = Invoke-Command -ComputerName $s -ScriptBlock $sbCommand -ArgumentList $s -Credential $ServerCredential -SessionOption (New-PSSessionOption -OperationTimeout 90000) #-ErrorAction SilentlyContinue
                     foreach ($res in $return)
                    {
                        $insert = InsertServerIISLoggingPath -CSHostname $res.ServerName -connect $res.Connect -AppPool $res.AppPool -SiteName $res.SiteName -State $res.State -LogEnabled $res.LogEnabled -LogPath $res.logPath -Binding $res.Binding -MaxLogFiles $res.MaxLogFiles -MaxLogFileSizeKB $res.MaxLogFileSizeKB -TraceFailedRequests $res.TraceFailedRequests -TraceFailedRequestsLogPath $res.TraceFailedRequestsLogPath -PhysicalPath $res.PhysicalPath

                    }
                }
                else
                {
                    [array]$return = [pscustomobject]@{
                                                            ServerName = $s
							                                Connect = 0
                                                            AppPool = "N/A"
                                                            SiteName = "N/A"
                                                            State = "N/A"
                                                            LogEnabled = "N/A"
                                                            LogPath = "ERROR:Failure to Connect on the Server. WinRM cannot complete the operation"
                                                            Binding = "N/A"
                                                            MaxLogFiles = "N/A"
                                                            MaxLogFileSizeKB = "N/A"
                                                            TraceFailedRequests = "N/A"
                                                            TraceFailedRequestsLogPath = "N/A"
															PhysicalPath="N/A"
                                                            StartDate = $startDate
                                                            FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
			                                                }
                                                             $insert = InsertServerIISLoggingPath -CSHostname $return.ServerName -connect $return.Connect -AppPool $return.AppPool -SiteName $return.SiteName -State $return.State -LogEnabled $return.LogEnabled -LogPath $return.logPath -Binding $return.Binding -MaxLogFiles $return.MaxLogFiles -MaxLogFileSizeKB $return.MaxLogFileSizeKB -TraceFailedRequests $return.TraceFailedRequests -TraceFailedRequestsLogPath $return.TraceFailedRequestsLogPath -PhysicalPath $return.PhysicalPath   
                }
              }catch
              {
                  $exception =  $_.Exception.Message; 
                  [array]$return = [pscustomobject]@{
                                                            ServerName = $s
							                                Connect = 0
                                                            AppPool = "N/A"
                                                            SiteName = "N/A"
                                                            State = "N/A"
                                                            LogEnabled = "N/A"
                                                            LogPath =  [String]$exception;
                                                            Binding = "N/A"
                                                            MaxLogFiles = "N/A"
                                                            MaxLogFileSizeKB = "N/A"
                                                            TraceFailedRequests = "N/A"
                                                            TraceFailedRequestsLogPath = "N/A"
															PhysicalPath="N/A"
                                                            StartDate = $startDate
                                                            FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
			                                         }
                                                      $insert = InsertServerIISLoggingPath -CSHostname $return.ServerName -connect $return.Connect -AppPool $return.AppPool -SiteName $return.SiteName -State $return.State -LogEnabled $return.LogEnabled -LogPath $return.logPath -Binding $return.Binding -MaxLogFiles $return.MaxLogFiles -MaxLogFileSizeKB $return.MaxLogFileSizeKB -TraceFailedRequests $return.TraceFailedRequests -TraceFailedRequestsLogPath $return.TraceFailedRequestsLogPath -PhysicalPath $return.PhysicalPath

                }
             }else
             {
                 [array]$return = [pscustomobject]@{
                                                            ServerName = $s
							                                Connect = 0
                                                            AppPool = "N/A"
                                                            SiteName = "N/A"
                                                            State = "N/A"
                                                            LogEnabled = "N/A"
                                                            LogPath =  "ERROR: $_.Exception.Message";
                                                            Binding = "N/A"
                                                            MaxLogFiles = "N/A"
                                                            MaxLogFileSizeKB = "N/A"
                                                            TraceFailedRequests = "N/A"
                                                            TraceFailedRequestsLogPath = "N/A"
															PhysicalPath="N/A"
                                                            StartDate = $startDate
                                                            FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
			                                         }
                                                      $insert = InsertServerIISLoggingPath -CSHostname $return.ServerName -connect $return.Connect -AppPool $return.AppPool -SiteName $return.SiteName -State $return.State -LogEnabled $return.LogEnabled -LogPath $return.logPath -Binding $return.Binding -MaxLogFiles $return.MaxLogFiles -MaxLogFileSizeKB $return.MaxLogFileSizeKB -TraceFailedRequests $return.TraceFailedRequests -TraceFailedRequestsLogPath $return.TraceFailedRequestsLogPath -PhysicalPath $return.PhysicalPath

                
             
             }

        } #endfor
    }else
    {
         WriteLog -msg "List of Servers is empty" -logname "ServerIISLoggingPath.log"
    }
    
    }catch
    {
          $exception =  $_.Exception.Message;
          [array]$return = [pscustomobject]@{
                                                            ServerName = $s
							                                Connect = 0
                                                            AppPool = "N/A"
                                                            SiteName = "N/A"
                                                            State = "N/A"
                                                            LogEnabled = "N/A"
                                                            LogPath =  [String]$exception
                                                            Binding = "N/A"
                                                            MaxLogFiles = "N/A"
                                                            MaxLogFileSizeKB = "N/A"
                                                            TraceFailedRequests = "N/A"
                                                            TraceFailedRequestsLogPath = "N/A"
															PhysicalPath="N/A"
                                                            StartDate = $startDate
                                                            FinishDate = [string]$(get-date -Format "yyyy-MM-dd hh:mm:ss")
			                                         }
                                                     $insert =  InsertServerIISLoggingPath -CSHostname $return.ServerName -connect $return.Connect -AppPool $return.AppPool -SiteName $return.SiteName -State $return.State -LogEnabled $return.LogEnabled -LogPath $return.logPath -Binding $return.Binding -MaxLogFiles $return.MaxLogFiles -MaxLogFileSizeKB $return.MaxLogFileSizeKB -TraceFailedRequests $return.TraceFailedRequests -TraceFailedRequestsLogPath $return.TraceFailedRequestsLogPath -PhysicalPath $return.PhysicalPath

    }
    
 
        #Sending E-mail
        SendEmail -logname "ServerIISLoggingPath.log"
        WriteLog -msg "Process has Finished." -logname "ServerIISLoggingPath.log"
 