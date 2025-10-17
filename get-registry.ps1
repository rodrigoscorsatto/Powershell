# Get Registry info based on parameters passed.

Param(
	$ComputerName,
	[switch]$Recursive,
	[int]$RecursionDepth,
	[string]$KeyPath,
	[string]$ValueName,
	[switch]$HKCR,
	[switch]$HKCU,
	[switch]$HKLM,
	[switch]$HKU,
	[switch]$HKCC,
	[switch]$HKDD
)
Begin
{	$script:CurrentErrorActionPreference = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"
	If($RecursionDepth){$Recursive = $True}
	$RegTypes = @{"REG_SZ" = 1; "REG_EXPAND_SZ" = 2; "REG_BINARY" = 3; "REG_DWORD" = 4; "REG_MULTI_SZ" = 7}
	$RegHives = @{"HKCR" = 2147483648; "HKCU" = 2147483649; "HKLM" = 2147483650; "HKU" = 2147483651;"HKCC" = 2147483653; "HKDD" = 2147483654}
	switch ($True)
	{
		$HKCR {$Hive = $RegHives.HKCR; break} #HKEY_CLASSES_ROOT
		$HKCU {$Hive = $RegHives.HKCU; break} #HKEY_CURRENT_USER
		$HKLM {$Hive = $RegHives.HKLM; break} #HKEY_LOCAL_MACHINE
		$HKU {$Hive = $RegHives.HKU; break} #HKEY_USERS
		$HKCC {$Hive = $RegHives.HKCC; break} #HKEY_CURRENT_CONFIG
		$HKDD {$Hive = $RegHives.HKDD; break} #HKEY_DYN_DATA
		default {$Hive = $RegHives.HKLM}
	}
	Function Get-TheRegistry
	{	Param(
			$ComputerName,
			$Hive,
			$KeyPath,
			$ValueName,
			$CurrentDepth
		)
		$CurrentDepth++
		Write-Progress -Activity "Get-Registry processing $ComputerName." -Status "Scanning $KeyPath" -ID 1
		Function Enum-RegKey
		{	Param($Reg,$Hive,$KeyPath)
			If(!$KeyPath){$KeyPath = ""}
			If($KeyPath.EndsWith("\")){$KeyPath = $KeyPath.TrimEnd("\")}
		 	$Keys = $reg.EnumKey($Hive,$KeyPath)
			If ($?)
			{	If($Keys.ReturnValue -eq 0)
				{	For($i=0; $i -le $Keys.sNames.Count–1; $i++)
					{
						$Object = New-Object psobject
						$Object | Add-Member NoteProperty KeyPath ""
						If($KeyPath -eq ""){$Object.KeyPath = "$($Keys.sNames[$i])"}
						Else{$Object.KeyPath = "$KeyPath\$($Keys.sNames[$i])"}
						$Object
					}
				}
			}
		}
		Function Enum-RegValues
		{	Param($Reg,$Hive,$KeyPath,$ValueName)
			Function New-Value($ValueName="",$ValueType="")
			{	$Object = New-Object psobject
				$Object | Add-Member NoteProperty ValueName $ValueName -PassThru | Add-Member NoteProperty ValueType $ValueType
				$Object
			}
			If($ValueName -eq "(Default)"){New-Value -ValueType $RegTypes.REG_SZ}
			Else
			{
			 	$Values = $reg.EnumValues($Hive,$KeyPath)
				If ($?)
				{	If($Values.ReturnValue -eq 0)
					{
						If($Values.sNames -eq $Null -And !$ValueName){New-Value -ValueType $RegTypes.REG_SZ}
						If($Values.sNames -ne $Null)
						{	If($ValueName)
							{	For($i=0; $i -le $Values.sNames.Count–1; $i++)
								{	If($Values.sNames[$i] -like $ValueName)
									{	New-Value -ValueName $Values.sNames[$i] -ValueType $Values.Types[$i]
									}
								}
							}
							Else
							{	For($i=0; $i -le $Values.sNames.Count–1; $i++)
								{New-Value -ValueName $Values.sNames[$i] -ValueType $Values.Types[$i]}
							}
						}
					}
				}
			}
		}
		Function Get-RegValue
		{	Param($Reg,$Hive,$KeyPath,$ValueName,$ValueType)
		 	switch($ValueType)
			{
				$RegTypes.REG_SZ				{$Result = $reg.GetStringValue($Hive,$KeyPath,$ValueName)}
				$RegTypes.REG_EXPAND_SZ	{$Result = $reg.GetExpandedStringValue($Hive,$KeyPath,$ValueName)}
				$RegTypes.REG_BINARY		{$Result = $reg.GetBinaryValue($Hive,$KeyPath,$ValueName)}
				$RegTypes.REG_DWORD			{$Result = $reg.GetDwordValue($Hive,$KeyPath,$ValueName)}
				$RegTypes.REG_MULTI_SZ	{$Result = $reg.GetMultiStringValue($Hive,$KeyPath,$ValueName)}
			}
			If($? -eq $False){$ReturnValue = "Error getting value :: $($Error[0].Exception.Message)"}
			ElseIf($Result.ReturnValue -ne 0){$ReturnValue = "Non-zero ReturnValue: $($Result.ReturnValue)"}
			Else
			{	If($ValueType -eq $RegTypes.REG_SZ -Or $ValueType -eq $RegTypes.REG_EXPAND_SZ -Or $ValueType -eq $RegTypes.REG_MULTI_SZ) {$ReturnValue = $Result.sValue}
				If($ValueType -eq $RegTypes.REG_DWORD){$ReturnValue = $Result.uValue}
				If($ValueType -eq $RegTypes.REG_BINARY)
				{	$ReturnValue = $Null
					ForEach($byte in $Result.uValue){$ReturnValue += "{0}" -f $byte.tostring("x2")}
					#ForEach($byte in $Result.uValue){$ReturnValue += "{0}" -f $byte.tostring("x")} #http://pshscripts.blogspot.fr/2011/09/get-wmiregistrybinaryvalueps1.html
				}
			}
			$ReturnValue
		}
		$reg = [wmiclass]"\\$ComputerName\root\default:StdRegprov"
		If ($? -eq $False)
		{	Write-Warning "Failed to connect to StdRegprov on $ComputerName : $($Error[0])"}
		Else
		{
			$RegValues = Enum-RegValues $Reg $Hive $KeyPath $ValueName
			If($RegValues)
			{	ForEach($RegValue in $RegValues)
				{
					$Object = New-Object psobject
					$Object | Add-Member NoteProperty ComputerName $ComputerName -PassThru | Add-Member NoteProperty Hive $Hive -PassThru |
						Add-Member NoteProperty KeyName "" -PassThru | Add-Member NoteProperty KeyPath $KeyPath -PassThru | 
						Add-Member NoteProperty ParentKeyName "" -PassThru | Add-Member NoteProperty ParentKeyPath "" -PassThru |
						Add-Member NoteProperty ValueName $RegValue.ValueName -PassThru |
						Add-Member NoteProperty ValueType $RegValue.ValueType -PassThru | 
						Add-Member NoteProperty Value (Get-RegValue $Reg $Hive $KeyPath $RegValue.ValueName $RegValue.ValueType)
					If($Object.ValueName -eq ""){$Object.ValueName = "(Default)"}
					[array]$KeyParts = $KeyPath.Split("\")
					$Object.KeyName = $KeyParts[$KeyParts.Count-1]
					If($KeyParts.Count -gt 1)
					{	$Object.ParentKeyName = $KeyParts[$KeyParts.Count-2]
						$Object.ParentKeyPath = $KeyParts[0..$($KeyParts.Count-2)] -join '\'
					}
					$Object.ValueType = ($RegTypes.GetEnumerator() | ?{$_.Value -eq $Object.ValueType}).Name
					$Object.Hive = ($RegHives.GetEnumerator() | ?{$_.Value -eq $Object.Hive}).Name
					$Object
				}
			}
			If($Recursive)
			{	If(!$RecursionDepth -Or $CurrentDepth -le $RecursionDepth)
				{$RegKeys = Enum-RegKey $Reg $Hive $KeyPath
					If($RegKeys)
					{	ForEach($RegKey in $RegKeys)
						{ Get-TheRegistry $ComputerName $Hive $RegKey.KeyPath $ValueName $CurrentDepth}
					}
				}
			}
		}
	}
}
Process
{
	If($ComputerName)
	{	ForEach($Computer in $ComputerName){Get-TheRegistry $Computer $Hive $KeyPath $ValueName}
	}
	Else
	{	Get-TheRegistry $_ $Hive $KeyPath $ValueName
	}
}