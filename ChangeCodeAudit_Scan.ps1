# Find accounts in code paths

$ToExecute = 
{
   param([string] $server, $PathList )
	
	try
	{

		$Results=@()
		
		foreach ($Path in $PathList)
		{
                if (($Path.Contains('\LOG')) -or ($Path.Contains('__LOG')) -or ($Path.Contains('__LOGS') -or ($Path.Contains('__LOGS') -or ($Path.Contains('DROPBOX')) -or ($Path.Contains('IMAGE'))-or ))){
                    $SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		    			    'CSHostname'  = $server;
		    			    'CodePath'=$Path;
                            'Files' = "N/A";
		    			    'Account'    = $Account;
						    'Status'    = "Warning";
						    'Exception'= "Path was ignored";
					  })
                 }
                 else
                 {
				    
                        if (Test-Path($Path))
                        {

                            $PathResult = Get-ChildItem -Path $Path -Include "*.bat","*.ps1","*.exe","*.config","*.json","*.txt" -Exclude "*.log" -Recurse|Select-String -List -Pattern $Account |Select-Object Path -ErrorAction SilentlyContinue
                            
				            $paths = $PathResult.Path
				
				            if ($paths)
				            {
					            $SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		    			            'CSHostname'  = $server;
		    			            'CodePath'=$Path;
                                    'Files' = $paths;
		    			            'Account'    = $Account;
						            'Status'    = "Found";
						            'Exception'= "N/A";
					            })				
				            }else
				            {
					            $SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		    			            'CSHostname'  = $server;
		    			            'CodePath'=$Path;
                                    'Files' = "Not Found";
		    			            'Account'    = $Account;
						            'Status'    = "Not Found";
						            'Exception'= "N/A";
					            })
				            }
                       }else
                       {
                         $SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		    			            'CSHostname'  = $server;
		    			            'CodePath'=$Path;
                                    'Files' = "Not Found";
		    			            'Account'    = $Account;
						            'Status'    = "Not Found";
						            'Exception'= "Path does not Exists!";
					            })

                       }
                }
				$Results+=$SouceCodeObj	
			}		
			
		
		#Write-Host $Results
		return $Results
	}catch
	{
		$SouceCodeObj = New-Object -TypeName PsObject -Property(@{
			'CSHostname'  = $server;
			'CodePath'=$Path;
            'Files' = "ERROR";
			'Account'    = $Account;
			'Status'    = "ERROR";
			'Exception'= $_.Exception.Message;
		})
		$Results+=$SouceCodeObj
		return $Results
	}
		
		
}


	 
	 #Remove @ and \ + domain 
	 Function Format_Account([String] $AccountList)
	 {
	 	$Accounts =  $AccountList -split(";")
		 [array]$result = @()
		foreach ($Account in $Accounts)
		{
		 	$len = $Account.Length-$Account.IndexOf('\')
			$len--
			$index=$Account.IndexOf('\')+1
			$str = $Account.Substring($index, $len)
		 
		 	$index=$Account.IndexOf('@')
			if ($index -gt 0){
				$str = $Account.Substring(0, $index)
			}
			$str = $str.ToUpper()
			$result+=$str
		}
		$result = $result | Select-Object -Unique 
		return $result
	 }
	 Function Format_Path([String] $PathList)
	 {
        $PathList = $PathList.Replace('N/A;','')
        $PathList = $PathList.Replace('N/A','')
	 	$Paths =  $PathList -split(";")
		[array]$result = @()
    
		foreach ($Path in $Paths)
		{
            if ($Path -like '*.EXE*')
            {
		 	    $Path = $Path.Replace('"','') # remove double quotes
			    $Path = $Path.ToUpper()
			    $index=  $Path.IndexOf('.EXE')
			    $Path = $Path.Substring(0,$index)
			    $Path = $Path |Split-Path #remove .exe path
            }
			$result+=$Path
            
		}
		$result = $result | Select-Object -Unique 
		return $result
	}
	 

Function FindAccount([String] $HostName, $PathList, $AccountList)
{
	if (Test-Connection -Cn $HostName -BufferSize 16 -Count 1 -ea 0 -quiet)
        {
			if (Test-WSMan -ComputerName $HostName)
            {
						$Error.Clear();
	    				$res = Invoke-command -computer $HostName -ArgumentList $HostName,$AccountList,$PathList -ScriptBlock $ToExecute -ErrorAction SilentlyContinue -credential $DBCredential -SessionOption (New-PSSessionOption -OperationTimeout 600000 -IdleTimeout 600000 ) #wait max 10 mins to process
						
						if ($Error.Count -gt 0)
						{
							$SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		                        'CSHostname'  = $HostName;
		                        'CodePath'= $PathList;
                                'Files' = "ERROR";
		                        'Account'    = $AccountList;
								'Status'    = "ERROR";
		                        'Exception'= $Error[0];
	                     		})
								$Insert = InsertData -CSHostname $SouceCodeObj.CSHostname -CodePath $SouceCodeObj.CodePath -files $SouceCodeObj.Files  -Account $SouceCodeObj.Account -Status $SouceCodeObj.Status -Exception $SouceCodeObj.Exception
							
						}
						else
						{
							foreach ($response in $res)
							{
                                $files = $response.Files;

                                foreach ($file in $files)
                                {
								    $SouceCodeObj = New-Object -TypeName PsObject -Property(@{
			                            'CSHostname'  = $response.CSHostName;
			                            'CodePath'= $response.CodePath;
                                        'Files' = $file;
			                            'Account'    = $response.Account;
									    'Status'    = $response.Status;
			                            'Exception'= $response.Exception;
		                     		    })
									    $Insert = InsertData -CSHostname $SouceCodeObj.CSHostname -CodePath $SouceCodeObj.CodePath -files $SouceCodeObj.Files -Account $SouceCodeObj.Account -Status $SouceCodeObj.Status -Exception $SouceCodeObj.Exception						
                                }
							}
						}
		
			}else
			{
				$SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		             'CSHostname'  = $HostName;
		             'CodePath'= $PathList;
                     'Files' = "ERROR"
		             'Account'    = $AccountList;
					 'Status'    = "ERROR";
		             'Exception'= "ERROR: The client cannot connect to the destination specified in the request";
	            })
				$Insert = InsertData -CSHostname $SouceCodeObj.CSHostname -CodePath $SouceCodeObj.CodePath -files $SouceCodeObj.Files -Account $SouceCodeObj.Account -Status $SouceCodeObj.Status -Exception $SouceCodeObj.Exception
			
			}	
	 }else
	 {
	 	$SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		    'CSHostname'  = $HostName;
		    'CodePath'= $PathList;
            'Files' = "ERROR"
		    'Account'    = $AccountList;
			'Status'    = "ERROR";
		    'Exception'= "ERROR: Server is not Reachable";
	    })
		$Insert = InsertData -CSHostname $SouceCodeObj.CSHostname -CodePath $SouceCodeObj.CodePath -files $SouceCodeObj.Files -Account $SouceCodeObj.Account  -Status $SouceCodeObj.Status -Exception $SouceCodeObj.Exception
	 
	 }




}
try
{	 
	 $Info = GetAccountPathInfo
     foreach ($list in $info)
	 {
	 	
		$AccountList=$list.accounts
		$AccountList = Format_Account -AccountList $AccountList
		$PathList = $list.paths
	 	$PathList = Format_Path -PathList $PathList
		$HostName = $list.CSHostname     
        WriteLog -msg "Processing Server: $HostName" -logname "SourcecodeAccountFinder.log"
        Write-Host "Processing Server: $HostName Accounts: $AccountList Paths: $PathList" 
		FindAccount -HostName $HostName -PathList $PathList -AccountList $AccountList
	 }
}catch
{
		$SouceCodeObj = New-Object -TypeName PsObject -Property(@{
		    'CSHostname'  = $HostName;
		    'CodePath'= $Path;
            'Files' = "ERROR"
		    'Account'    = $Account;
			'SourceType'    = "Unknown";
			'Status'    = "ERROR";
		    'Exception'= "ERROR: $_.Exception.Message";
	    })
		$Insert = InsertData -CSHostname $SouceCodeObj.CSHostname -CodePath $SouceCodeObj.CodePath -files $SouceCodeObj.Files -Account $SouceCodeObj.Account -SourceType $SouceCodeObj.SourceType -Status $SouceCodeObj.Status -Exception $SouceCodeObj.Exception
}

#Sending E-mail
SendEmail -logname "SourcecodeAccountFinder.log"
WriteLog -msg "Process has Finished." -logname "SourcecodeAccountFinder.log"