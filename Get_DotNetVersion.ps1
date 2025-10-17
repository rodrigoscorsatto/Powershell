<#
.DESCRIPTION
    1)	You can use this script to get dot net version details 


.NOTES
    File Name      : Get_DotNetVersion.ps1
    Author         : Rodrigo Scorsatto 
    Prerequisite   : PowerShell V2 or upper
    Version        : 1.0 - 01/21/2021

#>


$ToExecute =
{
     param([string] $server)  
    
   
   try{
        $key = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP"
        if (Test-Path $key)
        {
            $element = Get-ChildItem $key -recurse | Get-ItemProperty -name Version,Release -EA 0 | Where { $_.PSChildName -match '^(?!S)\p{L}'} |Select PSChildName, Version, Release
            $result = @()
		    foreach ($item in $element)
            {

                $Obj = New-Object -TypeName PsObject -Property(@{
                    'Server'  = $server;
                    'Name'= [String]$item.PSChildName;
                    'Version'=[String] $item.Version;
                    'Release'= [String]$item.Release;
                    'Exception'= "N/A";
                })
                $result += $Obj
            }
            return $result 
         }else
         {
             $Obj = New-Object -TypeName PsObject -Property(@{
                'Server'  = $server;
                'Name'= "N/A";
                'Version' = "N/A";
                'Release'    = "N/A";
                'Exception'= "Registry: $key do not exists";
            })

         }
         return $Obj
    }catch
    {
        $exception = $_.Exception.Message;
        $Obj = New-Object -TypeName PsObject -Property(@{
            'Server'  = $server;
            'Name'= "N/A";
            'Version' = "N/A";
            'Release'    = "N/A";
            'Exception'= $exception;
        })
           
        return $Obj
    
    }
}


    
     Function InsertDotNetVersionData([String] $CSHostname, [String]$Name, [String] $Version, [String]$Release, [String]$Exception)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Name: $Name, Version: $Version, Release: $Release, Exception: $Exception" -logname "GetdotNetVersion.log"

          $Query = "INSERT INTO DotNetVersion (CSHostname, Name,[Version],Release,Exception) VALUES('{0}','{1}','{2}','{3}','{4}')" -f $CSHostname,$Name, $Version, $Release, $Exception
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
            $exception = $_.Exception.Message;
            WriteLog -msg "There is an error when trying to insert data. Error: $exception" -logname "GetdotNetVersion.log"
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
           $delete = DeleteAllDataTable -tablename "DotNetVersion" -databasename "ServerAudit" -logname "GetdotNetVersion.log"
           for ($i=0; $i -lt $HostNames.Length; $i++)
           {
                $s = $HostNames[$i].CSHostname;
                WriteLog -msg "Processing Server: $s" -logname "GetdotNetVersion.log"
                if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet)
                {
                    if (Test-WSMan -ComputerName $s)
                    {
						$Error.Clear();
                        $res = Invoke-command -computer $s -ScriptBlock $ToExecute -ArgumentList $s -credential $ServerCredential -SessionOption (New-PSSessionOption -OperationTimeout 90000)
                        if ($Error.Count -gt 0)
						{
						 
						     $exception = $Error[0]
                             $Obj = New-Object -TypeName PsObject -Property(@{
                                'Server'  = $s;
                                'Name'= "N/A";
                                'Version' = "N/A";
                                'Release'    = "N/A";
                                'Exception'= $exception;
                             })
                             $insert = InsertDotNetVersionData -CSHostname $Obj.Server -Name $Obj.Name -Version $Obj.Version -Release $Obj.Release -Exception $Obj.Exception
      			               	
						}
						else
						{
							foreach ($result in $res)
	                        {
	                           
	                            $insert = InsertDotNetVersionData -CSHostname $result.Server -Name $result.Name -Version $result.Version -Release $result.Release -Exception $result.Exception
	                        }
						}
                     }else
                     {
                        
                        $Obj = New-Object -TypeName PsObject -Property(@{
                                'Server'  = $s;
                                'Name'= "N/A";
                                'Version' = "N/A";
                                'Release'    = "N/A";
                                'Exception'= "ERROR: The client cannot connect to the destination specified in the request";
                             })
                             $insert = InsertDotNetVersionData -CSHostname $Obj.Server -Name $Obj.Name -Version $Obj.Version -Release $Obj.Release -Exception $Obj.Exception
                     
                     }
                }else
                {
                    $exception = $_.Exception.Message
                    $Obj = New-Object -TypeName PsObject -Property(@{
                                'Server'  = $s;
                                'Name'= "N/A";
                                'Version' = "N/A";
                                'Release'    = "N/A";
                                'Exception'= $exception;
                             })
                             $insert = InsertDotNetVersionData -CSHostname $Obj.Server -Name $Obj.Name -Version $Obj.Version -Release $Obj.Release -Exception $Obj.Exception
                }
            }
            
        }
        else
        {
            WriteLog -msg "List of Servers is empty" -logname "GetdotNetVersion.log"
        }
    }catch
    {
           $exception = $_.Exception.Message
           $Obj = New-Object -TypeName PsObject -Property(@{
                  'Server'  = $s;
                  'Name'= "N/A";
                  'Version' = "N/A";
                  'Release'    = "N/A";
                  'Exception'= $exception;
            })
            $insert = InsertDotNetVersionData -CSHostname $Obj.Server -Name $Obj.Name -Version $Obj.Version -Release $Obj.Release -Exception $Obj.Exception
    }

        #Sending E-mail
        SendEmail -logname "GetdotNetVersion.log"
        WriteLog -msg "Process has Finished." -logname "GetdotNetVersion.log"