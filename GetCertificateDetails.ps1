# GET CERTIFICATE DETAILS FROM THE IIS BINDINGS


function GetCertInfo ([string] $strcomputer)
{
   try
        {
            Import-Module WEbAdministration 
            $Binding = Get-ChildItem -Path IIS:SSLBindings -ErrorAction Stop
            $CertList= @()
            foreach ($bind in $Binding)
            {
                if ($bind.Sites)
                {
                    $certificate = Get-ChildItem -Path CERT:LocalMachine/My -ErrorAction Stop
                    $found = 0;
                    foreach ($cert in $certificate)
                    {
                        if ($cert.Thumbprint -eq $bind.Thumbprint)
                        {
                            $result = New-Object -TypeName PsObject -Property(@{
                            'ComputerName' = [string] $strcomputer;
                            'Site' = [string] $bind.Sites.Value;
                            'Thumbprint' = $cert.Thumbprint;
                            'CertificateName'=$cert.Subject;
                            'DNSNameList'=$cert.DnsNameList;
                            'ExpireDate'=$cert.NotAfter;
                            'Issuer'=$cert.Issuer;
                         })  
                         $found=1;
                         $CertList+=$result
                        }
                    }
                    if ($found -eq 0)
                    {
                         $result = New-Object -TypeName PsObject -Property(@{
                            'ComputerName' = [string] $strcomputer;
                            'Site' = [string] $bind.Sites.Value;
                            'Thumbprint' = "Binding exists but no Certificate is assigned to it";
                            'CertificateName'="Binding exists but no Certificate is assigned to it";
                            'DNSNameList'="Binding exists but no Certificate is assigned to it";
                            'ExpireDate'="Binding exists but no Certificate is assigned to it";
                            'Issuer'="Binding exists but no Certificate is assigned to it";
                         })  
                         $CertList+=$result
                    }
                }
                
            }
            if ($CertList)
            {
                return $CertList
            }
            else
            {
                 $result = New-Object -TypeName PsObject -Property(@{
                            'ComputerName' = [string] $strcomputer;
                            'Site' = "N/A";
                            'Thumbprint' = "Not Found";
                            'CertificateName'="Not Found";
                            'DNSNameList'="Not Found";
                            'ExpireDate'="Not Found";
                            'Issuer'="Not Found";
                         })  
                         return $result

            }
        }catch
             {
                    $ErrorMessage = $_.Exception.Message

                    $result = New-Object -TypeName PsObject -Property(@{
                        'ComputerName' = [string] $strcomputer;
                        'Site' = "ERROR";
                        'Thumbprint' = "ERROR";
                        'CertificateName'="ERROR";
                        'DNSNameList'="ERROR";
                        'ExpireDate'="ERROR";
                        'Issuer'=$ErrorMessage;
                    })  
                    return $result
              }
    
}

# ************************************************* Main **********************************************




  

     Function InsertServerSSLCertificates([String] $CSHostname, [String]$Site, [String]$Thumbprint, [String]$CertificateName, [String]$DNSNameList, [String]$ExpireDate, [String]$Issuer)
     {
       try{
          WriteLog -msg "Inserting Server: $CSHostname with data: Site: $Site, Certificate: $CertificateName, Expire date: $ExpireDate" -logname "ServerSSLCertificates.log" 
		  $DNSNameList = $DNSNameList -replace("'",""); 
          $Query = "INSERT INTO SSLCertificates (CSHostname, [site], Thumbprint, CertificateName, DNSNameList, [ExpireDate], Issuer) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}')" -f $CSHostname, $Site, $Thumbprint, $CertificateName, $DNSNameList, $ExpireDate, $Issuer
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
            WriteLog -msg "There is an error when trying to insert data. Error: $exception" -logname "ServerSSLCertificates.log" 
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
        $delete = DeleteAllDataTable -tablename "SSLCertificates" -databasename "ServerAudit" -logname "ServerSSLCertificates.log" 
           for ($i=0; $i -lt $HostNames.Length; $i++)
           {
                $s = $HostNames[$i].CSHostname;   
              if (Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet)
              {

                if (Test-WSMan -ComputerName $s)
                {
                    $cmd = Invoke-command -computer $s -scriptBlock ${function:GetCertInfo} -ArgumentList $s -Credential $ServerCredential -SessionOption (New-PSSessionOption -OperationTimeout 90000)
                    foreach ($CertList in $cmd)
                    {
                       $insert = InsertServerSSLCertificates -CSHostname $CertList.ComputerName -Site $CertList.Site -Thumbprint $CertList.Thumbprint -CertificateName $CertList.CertificateName -DNSNameList $CertList.DNSNameList -ExpireDate $CertList.ExpireDate -Issuer $CertList.Issuer 
                    }
                }else
                {
                     $ErrorMessage = $_.Exception.Message
                     $result = New-Object -TypeName PsObject -Property(@{
                        'ComputerName' = [string] $s;
                        'Site' = "ERROR";
                        'Thumbprint' = "ERROR";
                        'CertificateName'="ERROR";
                        'DNSNameList'="ERROR";
                        'ExpireDate'="ERROR";
                        'Issuer'=[String]$ErrorMessage
                    })  
                   $insert = InsertServerSSLCertificates -CSHostname $result.ComputerName -Site $result.Site -Thumbprint $result.Thumbprint -CertificateName $result.CertificateName -DNSNameList $result.DNSNameList -ExpireDate $result.ExpireDate -Issuer $result.Issuer 
                }
              }else
              {
                $ErrorMessage = $_.Exception.Message
                     $result = New-Object -TypeName PsObject -Property(@{
                        'ComputerName' = [string] $s;
                        'Site' = "ERROR";
                        'Thumbprint' = "ERROR";
                        'CertificateName'="ERROR";
                        'DNSNameList'="ERROR";
                        'ExpireDate'="ERROR";
                        'Issuer'=[String]$ErrorMessage;
                    })  
                    $insert = InsertServerSSLCertificates -CSHostname $result.ComputerName -Site $result.Site -Thumbprint $result.Thumbprint -CertificateName $result.CertificateName -DNSNameList $result.DNSNameList -ExpireDate $result.ExpireDate -Issuer $result.Issuer 

              }
            }
    }else
    {
        WriteLog -msg "List of Servers is empty" -logname "ServerSSLCertificates.log" 
    }
 }catch
 {
    $ErrorMessage = $_.Exception.Message
                     $result = New-Object -TypeName PsObject -Property(@{
                        'ComputerName' = [string] $s;
                        'Site' = "ERROR";
                        'Thumbprint' = "ERROR";
                        'CertificateName'="ERROR";
                        'DNSNameList'="ERROR";
                        'ExpireDate'="ERROR";
                        'Issuer'= [String]$ErrorMessage;
                    })  
                    $insert = InsertServerSSLCertificates -CSHostname $result.ComputerName -Site $result.Site -Thumbprint $result.Thumbprint -CertificateName $result.CertificateName -DNSNameList $result.DNSNameList -ExpireDate $result.ExpireDate -Issuer $result.Issuer 

 }
 
  try
  {
	  #Sending E-mail
	  SendEmail -logname "ServerIISLoggingPath.log"
	  WriteLog -msg "Process has Finished."-logname "ServerSSLCertificates.log" 
  }catch
  { 
	   WriteLog -msg "ERROR: There is an Error when Trying to send the e-mail" -logname "ServerSSLCertificates.log" 
     $exception =  $_.Exception.Message;
	   WriteLog -msg $exception -logname "ServerSSLCertificates.log" 
  }