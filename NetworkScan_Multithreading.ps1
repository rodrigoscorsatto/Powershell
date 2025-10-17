<#

  Script to get network details from a list of servers using powershell jobs

.NOTES
    File Name      : .\NetworkScan_Multithreading.ps1
    Author         : Rodrigo Scorsatto 
    Prerequisite   : PowerShell V2 or upper
    Version        : 1.0 - 02/09/2021

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
    $nwINFO = Get-WmiObject -ComputerName $server Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null } 
    
         if ($?)
    {
                  
    
    $nwDescrip = $nwINFO.Description 
    $nwIPADDR = $nwINFO.IPAddress[0] 
    $nwSUBNET = $nwINFO.IpSubnet[0] 
    $nwGateWay = $nwINFO.DefaultIPGateway -join ',' 
    $nwMacADD = $nwINFO.MACAddress 
    $nwDNS = $nwINFO.DNSServerSearchOrder -join ',' 
    
     
    $results = New-Object -TypeName PsObject -Property(@{
      'Server Name' = $server;
      'Adapter' = $nwDescrip;
      'Ip Address'=$nwIPADDR;
      'Ip Subnet'=$nwSUBNET;
      'Default Ip Gateway'=$nwGateWay;
      'MAC Address'=$nwMacADD;
      'DNS Server'=$nwDNS;
      'Exception' = "N/A"
  })
  return $results
}else{
  
  $results = New-Object -TypeName PsObject -Property(@{
    'Server Name' = $server;
    'Adapter' = "ERROR";
    'Ip Address'="ERROR";
    'Ip Subnet'= "ERROR";
    'Default Ip Gateway'="ERROR";
    'MAC Address'="ERROR";
    'DNS Server'="ERROR";
    'Exception' = $_.Exception.Message;
})
return $results
         
}

   }catch{
                
    $results = New-Object -TypeName PsObject -Property(@{
      'Server Name' = $server;
      'Adapter' = "ERROR";
      'Ip Address'="ERROR";
      'Ip Subnet'= "ERROR";
      'Default Ip Gateway'="ERROR";
      'MAC Address'="ERROR";
      'DNS Server'="ERROR";
      'Exception' = $_.Exception.Message;
  })
  return $results
          
           
  }
}
}


     Function InsertNetworkScanData([String] $CSHostname, [String]$Adapter, [String]$MAC, [String]$IP_Address, [String]$Gateway, [String]$DNS, [String] $Subnet, [string]$Exception)
     {
       try{

          WriteLog -msg "Inserting Server: $CSHostname with data: Adapter: $Adapter, MAC: $MAC, IP Address: $IP_Address, Gateway IP:$Gateway, DNS: $DNS, Subnet IP: $Subnet, Exception: $Exception" -logname "NetworkScan.log"

          $Query = "INSERT INTO NetworkScan (CSHostname, [Adapter],[MAC_Address],[IP_Address],[Default_Gateway_IP],[DNS_Server],[Subnet_IP],[Exception]) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')" -f $CSHostname, $Adapter, $MAC, $IP_Address, $Gateway, $DNS, $Subnet, $Exception
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "NetworkScan.log"
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
               $delete =  DeleteAllDataTable -tablename "NetworkScan" -databasename "ServerAudit" -logname "NetworkScan.log"
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
                              'Server Name' = $s;
                              'Adapter' = "ERROR";
                              'Ip Address'="ERROR";
                              'Ip Subnet'= "ERROR";
                              'Default Ip Gateway'="ERROR";
                              'MAC Address'="ERROR";
                              'DNS Server'="ERROR";
                              'Exception' = "ERROR: Failure to Connect on the Server. WinRM cannot complete the operation";
                            }) 
                            InsertNetworkScanData -CSHostname $ServerInfo.'Server Name' -Adapter $ServerInfo.Adapter -MAC $ServerInfo.'MAC Address' -IP_Address $ServerInfo.'Ip Address' -Gateway $ServerInfo.'Default Ip Gateway' -DNS $ServerInfo.'DNS Server' -Subnet $ServerInfo.'Ip Subnet'-Exception $ServerInfo.Exception

                        }
                     }else
                     {

                      $ServerInfo = New-Object -TypeName PsObject -Property(@{
                        'Server Name' = $s;
                        'Adapter' = "ERROR";
                        'Ip Address'="ERROR";
                        'Ip Subnet'= "ERROR";
                        'Default Ip Gateway'="ERROR";
                        'MAC Address'="ERROR";
                        'DNS Server'="ERROR";
                        'Exception' = "ERROR: Server is not reachable";
                      }) 
                      InsertNetworkScanData -CSHostname $ServerInfo.'Server Name' -Adapter $ServerInfo.Adapter -MAC $ServerInfo.'MAC Address' -IP_Address $ServerInfo.'Ip Address' -Gateway $ServerInfo.'Default Ip Gateway' -DNS $ServerInfo.'DNS Server' -Subnet $ServerInfo.'Ip Subnet'-Exception $ServerInfo.Exception
                     }
               }
               Get-Job | Wait-Job -Timeout 180
               $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName     
               
                   
               foreach ($result in $results)
               {
                InsertNetworkScanData -CSHostname $result.'Server Name' -Adapter $result.Adapter -MAC $result.'MAC Address' -IP_Address $result.'Ip Address' -Gateway $result.'Default Ip Gateway' -DNS $result.'DNS Server' -Subnet $result.'Ip Subnet'-Exception $result.Exception
                 
               }
 
    }else
        {
                WriteLog -msg "There is an error when trying to get Server List from database" -logname "NetworkScan.log"
            }
}catch
      {
           
        $ServerInfo = New-Object -TypeName PsObject -Property(@{
          'Server Name' = $s;
          'Adapter' = "ERROR";
          'Ip Address'="ERROR";
          'Ip Subnet'= "ERROR";
          'Default Ip Gateway'="ERROR";
          'MAC Address'="ERROR";
          'DNS Server'="ERROR";
          'Exception' = $_.Exception.Message;
        }) 
        InsertNetworkScanData -CSHostname $ServerInfo.'Server Name' -Adapter $ServerInfo.Adapter -MAC $ServerInfo.'MAC Address' -IP_Address $ServerInfo.'Ip Address' -Gateway $ServerInfo.'Default Ip Gateway' -DNS $ServerInfo.'DNS Server' -Subnet $ServerInfo.'Ip Subnet'-Exception $ServerInfo.Exception
        
       }
        #Sending E-mail
        SendEmail -logname "NetworkScan.log"
        WriteLog -msg "Process has Finished." -logname "NetworkScan.log"
 
    
    }