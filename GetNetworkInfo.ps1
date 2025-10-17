# Get network info from a list o servers
param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage='Please inform the file path containing the Servers list.')][string] $HostNameList
)
$Result_path = "./Results.csv"
if ( $HostNameList)
{
    if (Test-Path $HostNameList)
    {
    
         if ( $HostNameList.Length -gt 0 ) 
         {
            $HostNames = Get-Content $HostNameList

            if (Test-Path $Result_path)
            {
            	Write-Host "Deleting Existing CSV Result file at" $Result_path  -ForegroundColor Yellow
	            Remove-Item $Result_path
            }
        
            foreach ($server in $HostNames)
            {
                
              #Write-Host "Processing Servers: $server"
              $nwINFO = Get-WmiObject -ComputerName $server Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null } 
              
              $nwServerName = $nwINFO.DNSHostName 
              $nwDescrip = $nwINFO.Description 
              $nwIPADDR = $nwINFO.IPAddress[0] 
              $nwSUBNET = $nwINFO.IpSubnet[0] 
              $nwGateWay = $nwINFO.DefaultIPGateway -join ',' 
              $nwMacADD = $nwINFO.MACAddress 
              $nwDNS = $nwINFO.DNSServerSearchOrder -join ',' 
              
              $results = New-Object -TypeName PsObject -Property(@{
                'Server Name' = $nwServerName;
                'Adapter' = $nwDescrip;
                'Ip Address'=$nwIPADDR;
                'Ip Subnet'=$nwSUBNET;
                'Default Ip Gateway'=$nwGateWay;
                'MAC Address'=$nwMacADD;
                'DNS Server'=$nwDNS;
            })
            Write-Host $results

              $results | Export-Csv -Path $Result_path -Append -NoTypeInformation

            }
            }else
		 {
			Write-Error ("List file $HostNameList is empty, Operation Canceled")
	     }
    }
  }else
	{
           Write-Error "Invalid Server Name List: $HostNameList.Operation Canceled"  
    }