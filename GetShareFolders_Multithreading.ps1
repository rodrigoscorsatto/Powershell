<#

  Get shared folders in the remote servers using Powershell JOBs

.NOTES
    File Name      : .\GetShareFolders_Multithreading.ps1
    Author         : Rodrigo Scorsatto 
    Prerequisite   : PowerShell V2 or upper
    Version        : 1.0 - 11/12/2018

REQUIREMENT: YOU MUST RUN THIS SCRIPT AS ADMINISTRATOR



#>


param (

    [Parameter(Mandatory=$false,ValueFromPipeline=$true,HelpMessage='Please inform the file path containing the Servers list.')][string] $HostNameList
 
);

begin
{
    
  
  $MaxThreads = 30 # Max concurrent running jobs.
  Get-Job | Remove-Job -Force
  
}
process
{
$exec = {
  param([string] $server)  
  process{
  
   try
   { 
        $Shares = get-WmiObject -class Win32_Share -computer $server -ErrorAction Stop
         if ($?)
    {
        $permissions = @()
        $result = @()
      
            ForEach ($element in $Shares)
            {   
                $ACL = $null
                $name = $element.name
                $objShareSec = Get-WMIObject -Class Win32_LogicalShareSecuritySetting -Filter "name='$name'"  -ComputerName $server 
                
                if ($objShareSec -ne $null)
                {
                    
                    $SD = $objShareSec.GetSecurityDescriptor().Descriptor    
                    foreach($ace in $SD.DACL)
                    {   
                        $permObj = @()
                        $UserName = $ace.Trustee.Name      
                        If ($ace.Trustee.Domain -ne $Null) {$UserName = "$($ace.Trustee.Domain)\$UserName"}    
                        If ($ace.Trustee.Name -eq $Null) {$UserName = $ace.Trustee.SIDString }      
                        [Array]$ACL += New-Object Security.AccessControl.FileSystemAccessRule($UserName, $ace.AccessMask, $ace.AceType)
                       
                    } 
                    
                    foreach ($perm in $ACL)
                    {
                        $user = $perm.IdentityReference
                        $access = $perm.FileSystemRights
                        $permObj = New-Object -TypeName PsObject -Property(@{
                        'ComputerName' = [string] $server;
                        'name'= $element.name;
                        'path'= $element.path;
                        'Description' =  $element.description;
                        'User' = $user;
                        'Access'= $access;
                        })
                       $permissions+= $permObj

                    }
                               
                }else
                {
                    $permObj = New-Object -TypeName PsObject -Property(@{
                        'ComputerName' = [string] $server;
                        'name'= $element.name;
                        'path'= $element.path;
                        'Description' =  $element.description;
                        'User' = "N/A";
                        'Access'= "N/A";
                        })
                       $permissions+= $permObj

                }
             
            }

    
            return $permissions

}else{
  $ServerInfo = New-Object -TypeName PsObject -Property(@{
                'ComputerName' = [string] $server;
                'name'= "ERROR";
                'path'= "ERROR";
                'Description' =  $_.Exception.Message;
                'User' = "ERROR";
                'Access'= "ERROR";
                
            })  
             $result+=$ServerInfo
             return $result
}

   }catch{
                
                $ServerInfo = New-Object -TypeName PsObject -Property(@{
                'ComputerName' = [string] $server;
                'name'= "ERROR";
                'path'= "ERROR";
                'Description' =  $_.Exception.Message;
                'AccessLevel' = "ERROR";
                'User' = "ERROR";
                'Access'= "ERROR";
                
            })  
             $result+=$ServerInfo
             return $result
           }
          
           
  }
}

     Function InsertShareFoldersData([String] $CSHostname, [String]$name, [String]$path, [String]$description, [String]$user, [String]$access)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Name: $name, Path: $path, Description: $description, User:$user, Permission: $access" -logname "ServerShareFolders_scan.log"

          $Query = "INSERT INTO ServerShareFolders (CSHostname, [name],[path],[description],[user],[access]) VALUES('{0}','{1}','{2}','{3}','{4}','{5}')" -f $CSHostname, $name, $path, $description, $user, $access
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "ServerShareFolders_scan.log"
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
               $delete =  DeleteAllDataTable -tablename "ServerShareFolders" -databasename "ServerAudit" -logname "ServerShareFolders_scan.log"
               $progressCounter=0
               for ($i=0; $i -lt $HostNames.Length; $i++)
               {
                    $s = $HostNames[$i].CSHostname;
                    if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet)
                    {
                        if (Test-WSMan -ComputerName $s)
                        {
                            $progressPercentage = [int]([float]($progressCounter / $HostNames.Count) * 100)
                            Write-Progress -Activity "Servers" -CurrentOperation "Processing server $s" -PercentComplete $progressPercentage -Status "$progressPercentage% Complete, Server Count: $progressCounter"
                            Start-Job -ScriptBlock $exec -ArgumentList $s -Name $s
                    
                            #Throttle jobs.
                            While($(Get-Job -State Running).Count -ge $MaxThreads) 
                            {
                                Get-Job | Wait-Job -Any -Timeout 180| Out-Null 
                            }
                            $progressCounter++
                        }else
                        {
                            $ServerInfo = New-Object -TypeName PsObject -Property(@{
                            'ComputerName' = $s;
                            'name'= "ERROR";
                            'path'= "ERROR";
                            'Description' = "ERROR: Failure to Connect on the Server. WinRM cannot complete the operation";
                            'AccessLevel' = "ERROR";
                            'User' = "ERROR";
                            'Access'= "ERROR";
                            }) 
                            InsertShareFoldersData -CSHostname $ServerInfo.ComputerName -name $ServerInfo.name -path $ServerInfo.path -description $ServerInfo.description -user $ServerInfo.user -access $ServerInfo.access

                        }
                     }else
                     {
                        $ServerInfo = New-Object -TypeName PsObject -Property(@{
                            'ComputerName' = [string] $s;
                            'name'= "ERROR";
                            'path'= "ERROR";
                            'Description' =  "ERROR: Server is not reachable";
                            'AccessLevel' = "ERROR";
                            'User' = "ERROR";
                            'Access'= "ERROR";
                            }) 
                            InsertShareFoldersData -CSHostname $ServerInfo.ComputerName -name $ServerInfo.name -path $ServerInfo.path -description $ServerInfo.description -user $ServerInfo.user -access $ServerInfo.access
                     }
               }
               Get-Job | Wait-Job -Timeout 180
               $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName     
               
                   
               foreach ($result in $results)
               {
                 InsertShareFoldersData -CSHostname $result.ComputerName -name $result.name -path $result.path -description $result.description -user $result.user -access $result.access
                 
               }
 
    }else
        {
                WriteLog -msg "There is an error when trying to get Server List from database" -logname "ServerShareFolders_scan.log"
            }
}catch
      {
           $ServerInfo = New-Object -TypeName PsObject -Property(@{
                  'ComputerName' = [string] $s;
                  'name'= "ERROR";
                  'path'= "ERROR";
                  'Description' = $_.Exception.Message;
                  'AccessLevel' = "ERROR";
                  'User' = "ERROR";
                  'Access'= "ERROR";
                  }) 
                  InsertShareFoldersData -CSHostname $ServerInfo.ComputerName -name $ServerInfo.name -path $ServerInfo.path -description $ServerInfo.description -user $ServerInfo.user -access $ServerInfo.access

       }
        #Sending E-mail
        SendEmail -logname "ServerShareFolders_scan.log"
        WriteLog -msg "Process has Finished." -logname "ServerShareFolders_scan.log"
 
    
    }