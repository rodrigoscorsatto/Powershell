# get service acocunt details from active directory using powershell jobs
begin
{
 $MaxThreads = 5 # Max concurrent running jobs.
 Get-Job | Remove-Job -Force
 Import-Module ActiveDirectory
 
}
process
{


 $exec = {
  param([string] $Account, [string] $Domain)  
  process
  {
    try
    {
      if (!$Domain -eq "N/A")
      {

          $dc = Get-ADDomainController -DomainName $Domain -Discover -NextClosestSite
          $dc_server=$dc.HostName[0]
      
	      $ad_user=Get-ADUser -Server $dc_server -Identity $Account -Properties *
      
          $AccountCreationDate=$ad_user.Created; 
            
          $AccountOwner = $ad_user.Description; 
      
          $LastBadPasswordAttempt = $ad_user.LastBadPasswordAttempt

          $LastLogonDate = $ad_user.LastLogonDate
      
          $LockedOut   = $ad_user.LockedOut
      
          $LockedoutDate = $ad_user.AccountLockoutTime

          $PasswordExpired = $ad_user.PasswordExpired
          $PasswordLastSet = $ad_user.PasswordLastSet
          $AccountStatus = $ad_user.Enabled


	        $AccountObj = New-Object -TypeName PsObject -Property(@{
                        'Account'  = $Account;
                        'Domain'= $Domain;
                        'DomainController' = $dc_server;
                        'Enabled' = $AccountStatus;
                        'CreationDate' =   $AccountCreationDate;
                        'AccountOwner'=  $AccountOwner;
                        'LastBadPasswordDate'=  $LastBadPasswordAttempt;
                        'LastLogonDate' =$LastLogonDate;
                        'LockedOut' =  $LockedOut;
					    'LockedoutDate' = $LockedoutDate;
					    'PasswordExpired' = $PasswordExpired;
					    'PasswordLastSetDate' = $PasswordLastSet;
					    'Exception' = 'N/A';
                   
                         })
                         return $AccountObj
    }else
    {
        try
        {
            $domain_try="amer.com"
            $dc = Get-ADDomainController -DomainName $domain_try -Discover -NextClosestSite
            $dc_server=$dc.HostName[0]
            $ad_user=Get-ADUser -Server $dc_server -Identity $Account -Properties *
            
            $AccountCreationDate=$ad_user.Created; 
            $AccountOwner = $ad_user.Description; 
            $LastBadPasswordAttempt = $ad_user.LastBadPasswordAttempt
            $LastLogonDate = $ad_user.LastLogonDate  
            $LockedOut   = $ad_user.LockedOut
            $LockedoutDate = $ad_user.AccountLockoutTime
            $PasswordExpired = $ad_user.PasswordExpired
            $PasswordLastSet = $ad_user.PasswordLastSet
            $AccountStatus = $ad_user.Enabled

	        $AccountObj = New-Object -TypeName PsObject -Property(@{
                        'Account'  = $Account;
                        'Domain'= $domain_try;
                        'DomainController' = $dc_server;
                        'Enabled' = $AccountStatus;
                        'CreationDate' =   $AccountCreationDate;
                        'AccountOwner'=  $AccountOwner;
                        'LastBadPasswordDate'=  $LastBadPasswordAttempt;
                        'LastLogonDate' =$LastLogonDate;
                        'LockedOut' =  $LockedOut;
					    'LockedoutDate' = $LockedoutDate;
					    'PasswordExpired' = $PasswordExpired;
					    'PasswordLastSetDate' = $PasswordLastSet;
					    'Exception' = 'N/A';
                   
                         })
                         return $AccountObj
        }catch
        {
            $ErrorMessage = $_.Exception.Message
            if ($ErrorMessage -match "Cannot find an object with identity")
            { 
                try
                {
                    $domain_try="apac.com"
                    $dc = Get-ADDomainController -DomainName $domain_try -Discover -NextClosestSite
                    $dc_server=$dc.HostName[0]
                    $ad_user=Get-ADUser -Server $dc_server -Identity $Account -Properties *
                    $AccountCreationDate=$ad_user.Created; 
                    $AccountOwner = $ad_user.Description;    
                    $LastBadPasswordAttempt = $ad_user.LastBadPasswordAttempt
                    $LastLogonDate = $ad_user.LastLogonDate
                    $LockedOut   = $ad_user.LockedOut
                    $LockedoutDate = $ad_user.AccountLockoutTime
                    $PasswordExpired = $ad_user.PasswordExpired
                    $PasswordLastSet = $ad_user.PasswordLastSet
                    $AccountStatus = $ad_user.Enabled

	                $AccountObj = New-Object -TypeName PsObject -Property(@{
                                    'Account'  = $Account;
                                    'Domain'= $domain_try;
                                    'DomainController' = $dc_server;
                                    'Enabled' = $AccountStatus;
                                    'CreationDate' =   $AccountCreationDate;
                                    'AccountOwner'=  $AccountOwner;
                                    'LastBadPasswordDate'=  $LastBadPasswordAttempt;
                                    'LastLogonDate' =$LastLogonDate;
                                    'LockedOut' =  $LockedOut;
					                'LockedoutDate' = $LockedoutDate;
					                'PasswordExpired' = $PasswordExpired;
					                'PasswordLastSetDate' = $PasswordLastSet;
					                'Exception' = 'N/A';
                   
                     })
                     return $AccountObj

                }catch
                {
                    $ErrorMessage = $_.Exception.Message
                    if ($ErrorMessage -match "Cannot find an object with identity"){ 
                        try
                        {
                            $domain_try="emea.com"
                            $dc = Get-ADDomainController -DomainName $domain_try -Discover -NextClosestSite
                            $dc_server=$dc.HostName[0]
                            $ad_user=Get-ADUser -Server $dc_server -Identity $Account -Properties *
                            $AccountCreationDate=$ad_user.Created; 
                            $AccountOwner = $ad_user.Description; 
                            $LastBadPasswordAttempt = $ad_user.LastBadPasswordAttempt
                            $LastLogonDate = $ad_user.LastLogonDate
                            $LockedOut   = $ad_user.LockedOut
                            $LockedoutDate = $ad_user.AccountLockoutTime
                            $PasswordExpired = $ad_user.PasswordExpired
                            $PasswordLastSet = $ad_user.PasswordLastSet
                            $AccountStatus = $ad_user.Enabled

	                        $AccountObj = New-Object -TypeName PsObject -Property(@{
                                            'Account'  = $Account;
                                            'Domain'= $domain_try;
                                            'DomainController' = $dc_server;
                                            'Enabled' = $AccountStatus;
                                            'CreationDate' =   $AccountCreationDate;
                                            'AccountOwner'=  $AccountOwner;
                                            'LastBadPasswordDate'=  $LastBadPasswordAttempt;
                                            'LastLogonDate' =$LastLogonDate;
                                            'LockedOut' =  $LockedOut;
					                        'LockedoutDate' = $LockedoutDate;
					                        'PasswordExpired' = $PasswordExpired;
					                        'PasswordLastSetDate' = $PasswordLastSet;
					                        'Exception' = 'N/A';
                   
                            })
                            return $AccountObj

                        }catch
                        {
                            
                           try
                        {
                            $domain_try="japn.com"
                            $dc = Get-ADDomainController -DomainName $domain_try -Discover -NextClosestSite
                            $dc_server=$dc.HostName[0]
                            $ad_user=Get-ADUser -Server $dc_server -Identity $Account -Properties *
                            $AccountCreationDate=$ad_user.Created; 
                            $AccountOwner = $ad_user.Description; 
                            $LastBadPasswordAttempt = $ad_user.LastBadPasswordAttempt
                            $LastLogonDate = $ad_user.LastLogonDate
                            $LockedOut   = $ad_user.LockedOut
                            $LockedoutDate = $ad_user.AccountLockoutTime
                            $PasswordExpired = $ad_user.PasswordExpired
                            $PasswordLastSet = $ad_user.PasswordLastSet
                            $AccountStatus = $ad_user.Enabled

	                        $AccountObj = New-Object -TypeName PsObject -Property(@{
                                            'Account'  = $Account;
                                            'Domain'= $domain_try;
                                            'DomainController' = $dc_server;
                                            'Enabled' = $AccountStatus;
                                            'CreationDate' =   $AccountCreationDate;
                                            'AccountOwner'=  $AccountOwner;
                                            'LastBadPasswordDate'=  $LastBadPasswordAttempt;
                                            'LastLogonDate' =$LastLogonDate;
                                            'LockedOut' =  $LockedOut;
					                        'LockedoutDate' = $LockedoutDate;
					                        'PasswordExpired' = $PasswordExpired;
					                        'PasswordLastSetDate' = $PasswordLastSet;
					                        'Exception' = 'N/A';
                   
                            })
                            return $AccountObj

                        }catch
                        {
                            
                            if ($ErrorMessage -match "Cannot find an object with identity")
                            {
                                
                                $AccountObj = New-Object -TypeName PsObject -Property(@{
                                    'Account'  = $Account;
                                    'Domain'= $Domain;
                                    'DomainController' = "Not Found"
                                    'Enabled' = "Not Found";
                                    'CreationDate' =   $null;
                                    'AccountOwner'=  "Not Found"
                                    'LastBadPasswordDate'=  $null;
                                    'LastLogonDate' = $null;
                                    'LockedOut' =  "Not Found"
					                'LockedoutDate' = $null;
					                'PasswordExpired' = "Not Found"
					                'PasswordLastSetDate' = $null;
					                'Exception' = "Not Found"
                   
                                 })
                                 return $AccountObj
                            }
                        }
                        }

                    }

                }
            }
          }

    }


    }catch
    {
         $AccountObj = New-Object -TypeName PsObject -Property(@{
                    'Account'  = $Account;
                    'Domain'= $Domain;
                    'DomainController' = $dc_server;
                    'Enabled' = "ERROR";
                    'CreationDate' =   $null;
                    'AccountOwner'=   "ERROR";
                    'LastBadPasswordDate'=   $null;
                    'LastLogonDate' =  $null;
                    'LockedOut' =   "ERROR";
					'LockedoutDate' =  $null;
					'PasswordExpired' =  $null;
					'PasswordLastSetDate' =  $null;
					'Exception' = $_.Exception.Message;
                   
                     })
                     return $AccountObj

  }
  }
}


   
    

   Function InsertServiceAccountData([String] $Account, [String]$Domain, [String] $DomainController,[String]$Enabled, [String]$CreationDate, [String]$AccountOwner, [String]$LastBadPasswordDate, [String]$LastLogonDate,[String]$LockedOut,[String]$LockedoutDate ,[String]$PasswordExpired,[String] $PasswordLastSetDate ,[String] $Exception)
     {
       try{

          WriteLog -msg "Inserting Account: $Account with data: Domain: $Domain, DomainController,$DomainController, Enabled:$Enabled CreationDate: $CreationDate, AccountOwner: $AccountOwner, LastBadPasswordDate: $LastBadPasswordDate,, LastLogonDate: $LastLogonDate, LockedOut: $LockedOut, PasswordExpired: $PasswordExpired, PasswordLastSetDate: $PasswordLastSetDate, Exception: $Exception" -logname "ServiceAccount.log"

          $Query = "INSERT INTO ServiceAccount (Account, Domain,DomainController,Enabled, CreationDate,AccountOwner,LastBadPasswordDate,LastLogonDate,LockedOut,LockedoutDate,PasswordExpired,PasswordLastSetDate,Exception) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}')" -f $Account,$Domain,$DomainController,$Enabled, $CreationDate,$AccountOwner,$LastBadPasswordDate,$LastLogonDate,$LockedOut,$LockedoutDate,$PasswordExpired,$PasswordLastSetDate,$Exception
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
            WriteLog -msg "There is an error when trying to insert data. Error: $_.Exception.Message" -logname "ServiceAccount.log"
            $result=$null
          }
          finally
          {
            $conn.close()

          }
     }
	 
     Function GetAccountList()
     {
       $table = $null 
       try{
              
			  $connectionString = GetConnectionString -DatabaseName "ServerAudit"
			  $connection = New-Object System.Data.SqlClient.SqlConnection
              $Query = GetAccountListServiceAccountDetailsQuery
              $connection.ConnectionString = $connectionString
              $connection.Open()
              $command = $connection.CreateCommand()
              $command.CommandText  = $query
              $result = $command.ExecuteReader()
              $table = new-object "System.Data.DataTable"
              $table.Load($result)
               
              return $table
          }catch
          {
            WriteLog -msg "There is an error when trying to get server data. Error: $_.Exception.Message" -logname "ServiceAccount.log"
            $table=$null
          }
          finally
          {
            $connection.close()

          }
     }
      #Remove @ and \ + domain like Americas
	 Function Format_Account([String] $AccountList)
	 {
	 	
		 [array]$result = @()
		foreach ($Account in $Accounts)
		{
		 	
            if (!(($account.Contains("\")) -or ($account.Contains("@"))))
            {
               $str=$account
               $domain="N/A"
            }
            else
            {
                $domain,$str = $Account.Split("\")

		 	    $index=$Account.IndexOf('@')
			    if ($index -gt 0){
                    $str,$domain = $Account.Split("@")
			    }
            }			

            $AccountObj = New-Object -TypeName PsObject -Property(@{
		                        'Account'  = $str;
		                        'Domain'= $domain;
	                     		})
			$result+=$AccountObj
		}
		$result = $result | Select-Object Account, Domain -Unique 
		return $result
	 }

try{
    	 $info = GetAccountList
         if ($info -ne $null)
         { 
               $delete = DeleteAllDataTable -tablename "ServiceAccount" -databasename "ServerAudit" -logname "ServiceAccount.log"
              
                   $accounts=$info.accounts 
                   $Accounts =  $accounts -split(";")
                   [array]$result = @()
                   foreach ($account in $Accounts )
                   {
            
		                $result+=$account
		           }
		           $result = $result | Select-Object -Unique 

                   
                   $AccountList = Format_Account -AccountList $result
                
                   foreach ($list in $AccountList)
                   {
                       $Account = $list.Account
                       $Domain = $list.Domain
                       
                       $progressPercentage = [int]([float]($progressCounter / $AccountList.Length) * 100)
                       Write-Progress -Activity "Account" -CurrentOperation "Processing Account $Account" -PercentComplete $progressPercentage -Status "$progressPercentage% Complete, Account Count: $progressCounter"
                       Start-Job -ScriptBlock $exec -ArgumentList $Account, $Domain -Name $Account
                    
                       #Throttle jobs.
                       While($(Get-Job -State Running).Count -ge $MaxThreads) 
                       {
                           #Write-host "Waiting Threads to release..." -ForegroundColor Yellow
                           Get-Job | Wait-Job -Any -Timeout 180| Out-Null 
                       }
                       $progressCounter++
                    }
                
            }
            else
            {
                WriteLog -msg "List of Accounts is empty" -logname "ServiceAccount.log"
            }
    }catch
    {
       
         $AccountObj = New-Object -TypeName PsObject -Property(@{
                    'Account'  = $Account;
                    'Domain'= $Domain;
                    'DomainController' =  "ERROR";
                    'Enabled' = "ERROR"
                    'CreationDate' =   $null;
                    'AccountOwner'=   "ERROR";
                    'LastBadPasswordDate'=   $null;
                    'LastLogonDate' =  $null;
                    'LockedOut' =   "ERROR";
					'LockedoutDate' =  $null;
					'PasswordExpired' =  "ERROR";
					'PasswordLastSetDate' =  $null;
					'Exception' = $_.Exception.Message;
                   
                     })
					 $insert = InsertServiceAccountData -Account $AccountObj.Account -Domain $AccountObj.Domain -DomainController $AccountObj.DomainController -Enabled $AccountObj.Enabled -CreationDate $AccountObj.CreationDate -AccountOwner $AccountObj.AccountOwner -LastBadPasswordDate $AccountObj.LastBadPasswordDate -LastLogonDate $AccountObj.LastLogonDate -LockedOut $AccountObj.LockedOut -LockedoutDate $AccountObj.LockedoutDate -PasswordExpired $AccountObj.PasswordExpired -PasswordLastSetDate $AccountObj.PasswordLastSetDate -Exception $AccountObj.Exception
    }
    Get-Job | Wait-Job -Timeout 180
    $results = Get-Job | Receive-Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName,PSShowComputerName 
    
    foreach ($r in $results){

         $AccountObj = New-Object -TypeName PsObject -Property(@{
                    'Account'  = $r.Account;
                    'Domain'= $r.Domain;
                    'DomainController' =  $r.DomainController;
                    'Enabled' = $r.Enabled;
                    'CreationDate' =   $r.CreationDate;
                    'AccountOwner'=   $r.AccountOwner;
                    'LastBadPasswordDate'= $r.LastBadPasswordDate;
                    'LastLogonDate' = $r.LastLogonDate;
                    'LockedOut' =   $r.LockedOut;
					'LockedoutDate' =  $r.LockedoutDate;
					'PasswordExpired' =  $r.PasswordExpired;
					'PasswordLastSetDate' =  $r.PasswordLastSetDate;
					'Exception' = $r.Exception;
                   
                     })
        $insert = InsertServiceAccountData -Account $AccountObj.Account -Domain $AccountObj.Domain -DomainController $AccountObj.DomainController -Enabled $AccountObj.Enabled -CreationDate $AccountObj.CreationDate -AccountOwner $AccountObj.AccountOwner -LastBadPasswordDate $AccountObj.LastBadPasswordDate -LastLogonDate $AccountObj.LastLogonDate -LockedOut $AccountObj.LockedOut -LockedoutDate $AccountObj.LockedoutDate -PasswordExpired $AccountObj.PasswordExpired -PasswordLastSetDate $AccountObj.PasswordLastSetDate -Exception $AccountObj.Exception
    }
     
        #Sending E-mail
        SendEmail -logname "ServiceAccount.log"
        WriteLog -msg "Process has Finished." -logname "ServiceAccount.log"
}

