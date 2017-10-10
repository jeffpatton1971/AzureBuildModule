function New-ResourceGroup
{
	param
	(
		[object]$ResourceGroupData,
		[object]$SubscriptionData
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		foreach ($ResourceGroup in $ResourceGroupData)
		{
			if ($ResourceGroup.ResourceGroupName -ne $null)
			{
				if (Get-AzureRmResourceGroup -Name $ResourceGroup.ResourceGroupName -ErrorAction SilentlyContinue)
				{
					Write-Host "Resource Group: $($ResourceGroup.ResourceGroupName) in $($ResourceGroup.Location) already exists" -ForegroundColor Yellow
				}
				else
				{
					Write-Host "Creating Resource Group: $($ResourceGroup.ResourceGroupName) in $($ResourceGroup.Location)" -ForegroundColor Green
					$Tag = @{ 
						BuildBy=$SubscriptionData.'Build By';
						BuildDate=$SubscriptionData.'Build Date';
						Ticket=$SubscriptionData.Ticket;
						Environment=$ResourceGroup.Environment}
					$status = New-AzureRmResourceGroup -Name $ResourceGroup.ResourceGroupName -Location $ResourceGroup.Location -Tag $Tag
					if ($status.ProvisioningState -eq 'Succeeded')
					{
						Write-Host "Success: Resource Group: $($ResourceGroup.ResourceGroupName) in $($ResourceGroup.Location)" -ForegroundColor Green
					}
					else
					{
						throw "Warning: Resource Group: $($ResourceGroup.ResourceGroupName) in $($ResourceGroup.Location) not in a Succeeded state, please validate"
					}
				}
			}
		}
	}
	catch
	{
		throw $_;
	}
}

function New-OmsWorkspace
{
	param
	(
		[object]$OmsWorkspaceData,
		[object]$SubscriptionData
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		foreach ($OmsWorkspace in $OmsWorkspaceData)
		{
			if($OmsWorkspace.OMSworkspaceName -ne $null)
			{
				$ResourceGroupName = $OmsWorkspace.OMSRSG
				$Template = $OmsWorkspace.Template
				$SAS = $OmsWorkspace.SAS

				$TemplateParameterObject = ConvertTo-Hashtable -PsObject $OmsWorkspace -Exclusionlist @('OMSRSG','Template','SAS');
				$status = New-AzureRmResourceGroupDeployment -Name $SubscriptionData.'Deployment Name' -ResourceGroupName $ResourceGroupName `
					-Mode Incremental `
					-TemplateParameterObject $TemplateParameterObject `
					-TemplateFile ("$template" + "$SAS") `
					-Force;
				if($status.ProvisioningState -eq 'Succeeded')
				{
					Write-Host "Success: Creating OMS Workspace: $($OmsWorkspaceData.OMSworkspaceName) in $ResourceGroupName" -ForegroundColor Green
				}
				else
				{
					throw "Warning: Creating OMS Workspace: $($OmsWorkspaceData.OMSworkspaceName) in $ResourceGroupName is not in a Succeeded state, please validate"
				}
			}
		}
	}
	catch
	{
		throw $_;
	}
}

function New-WebApp
{
	param
	(
		[object]$WebAppData,
		[object]$SubscriptionData
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		foreach ($WebApp in $WebAppData)
		{
			if($WebApp.webAppNames -ne $null)
			{
				$ResourceGroupName = $WebApp.WebAppRSG
				$Template = $WebApp.Template
				$SAS = $WebApp.SAS

				$TemplateParameterObject = ConvertTo-Hashtable -PsObject $WebApp -Exclusionlist @('WebAppRSG','Template','SAS');
				$status = New-AzureRmResourceGroupDeployment -Name $SubscriptionData.'Deployment Name' -ResourceGroupName $ResourceGroupName `
					-Mode Incremental `
					-TemplateParameterObject $TemplateParameterObject `
					-TemplateFile ("$template" + "$SAS") `
					-Force;
				if($status.ProvisioningState -eq 'Succeeded')
				{
					Write-Host "Success: Creating Web APP: $($WebApp.webAppNames) in $ResourceGroupName" -ForegroundColor Green
				}
				else
				{
					throw "Warning: Creating Web APP: $($WebApp.webAppNames) in $ResourceGroupName is not in a Succeeded state, please validate"
				}
			}
		}
	}
	catch
	{
		throw $_;
	}
}

function New-TrafficManager
{
	param
	(
		[object]$TrafficManagerData,
		[object]$SubscriptionData
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		foreach ($TrafficManager in $TrafficManagerData)
		{
			if ($TrafficManager.Name -ne $null)
			{
				if(Get-AzureRmTrafficManagerProfile -Name $TrafficManager.name -ResourceGroupName $TrafficManager.ResourceGroup -ErrorAction SilentlyContinue)
				{
					Write-Host "Traffic Manager Profile: $($TrafficManager.name) in $($TrafficManager.ResourceGroup) already exists" -ForegroundColor Yellow
				}
				else
				{
					Write-Host "Creating Traffic Manager: $($TrafficManager.name) in $($TrafficManager.ResourceGroup)" -ForegroundColor Green
					$Tag = @{ 
						BuildBy=$SubscriptionData.'Build By';
						BuildDate=$SubscriptionData.'Build Date';
						Ticket=$SubscriptionData.Ticket;
						Environment=$TrafficManager.Environment}
					$status = New-AzureRmTrafficManagerProfile -Name $TrafficManager.name -ResourceGroupName $TrafficManager.ResourceGroup -TrafficRoutingMethod $TrafficManager.trafficRoutingMethod -RelativeDnsName $TrafficManager.relativeName `
							-Ttl $TrafficManager.ttl -MonitorProtocol $TrafficManager.MonitorProtocol -MonitorPort $TrafficManager.MonitorPort -MonitorPath $TrafficManager.MonitorPath -Tag $Tag;
					if($status.ProfileStatus -eq 'Enabled')
					{
						Write-Host "Success: Creating Traffic Manager: $($TrafficManager.name) in $($TrafficManager.ResourceGroup)" -ForegroundColor Green
					}
					else
					{
						throw "Warning: Creating Traffic Manager: $($TrafficManager.name) in $($TrafficManager.ResourceGroup) is not in a Succeeded state, please validate"
					}
				}
			}
		}
	}
	catch
	{
		throw $_;
	}
}

function New-StorageAccount
{
	param
	(
		[object]$StorageAccountData,
		[object]$SubscriptionData
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		foreach ($StorageAccount in $StorageAccountData)
		{
			if ($StorageAccount.Name -ne $null)
			{
				if (Get-AzureRmStorageAccount -Name $StorageAccount.name -ResourceGroupName $StorageAccount.ResourceGroupName -ErrorAction SilentlyContinue)
				{
					Write-Host "Storage Account: $($StorageAccount.name) in $($StorageAccount.ResourceGroupName) already exists" -ForegroundColor Yellow
				}
				else
				{
					if (Get-AzureRmStorageAccountNameAvailability -Name $StorageAccount.name)
					{
						Write-Host "Creating Storage Account: $($StorageAccount.name) in $($StorageAccount.ResourceGroupName)" -ForegroundColor Green
						$Tag = @{ 
							BuildBy=$SubscriptionData.'Build By';
							BuildDate=$SubscriptionData.'Build Date';
							Ticket=$SubscriptionData.Ticket;
							Environment=$StorageAccount.Environment}
						$status = New-AzureRmStorageAccount -ResourceGroupName $StorageAccount.ResourceGroupName -Name $StorageAccount.name -SkuName $StorageAccount.SkuName -Location $StorageAccount.Location `
								-Kind $StorageAccount.kind -AccessTier $StorageAccount.AccessTier -EnableEncryptionService $StorageAccount.EnableEncryptionService -Tag $Tag
						if($status.ProvisioningState -eq 'Succeeded')
						{
							Write-Host "Success: Creating Storage Account: $($StorageAccount.name) in $($StorageAccount.ResourceGroupName)" -ForegroundColor Green
						}
						else
						{
							throw "Warning: Creating Storage Account: $($StorageAccount.name) in $($StorageAccount.ResourceGroupName) is not in a Succeeded state, please validate"
						}
					}
					else
					{
						throw "Error: Creating Storage Account: $($StorageAccount.name) is already in use"
					}
				}
			}
		}
	}
	catch
	{
		throw $_;
	}
}

function New-VirtualNetwork
{
	param
	(
		[object]$VirtualNetworkData,
		[object]$SubscriptionData
	)
	try 
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		foreach ($Vnet in $VirtualNetworkData)
		{
			if ($Vnet.Name -ne $null)
			{
				$ResourceGroupName = $Vnet.ResourceGroupName;
				$VNETDNS += @{Name=$VNET.Name;vnetRSG=$VNET.ResourceGroupName;PrimaryDNS=$VNET.PrimaryDNS;SecondaryDNS=$VNET.SecondaryDNS}
				$Template = $Vnet.Template
				$SAS = $Vnet.SAS
				$Size = $Vnet.VNETSize
				
				switch ($Size)
				{
					'Small'
					{
						$Vnet = ConvertTo-Hashtable -PsObject $Vnet -Exclusionlist 'PrimaryDNS','SecondaryDNS','ResourceGroupName','Template','SAS','vnetSize','environmentB','subnetDMZCIDRB','subnetAPPCIDRB','subnetINSCIDRB','subnetADCIDRB','subnetAGWCIDRB','environmentC','subnetDMZCIDRC','subnetAPPCIDRC','subnetINSCIDRC','subnetADCIDRC','subnetAGWCIDRC'
					}
					'Medium'
					{
							$Vnet = ConvertTo-Hashtable -PsObject $Vnet -Exclusionlist 'PrimaryDNS','SecondaryDNS','ResourceGroupName','Template','SAS','vnetSize','environmentC','subnetDMZCIDRC','subnetAPPCIDRC','subnetINSCIDRC','subnetADCIDRC','subnetAGWCIDRC'
					}
				}
				
				Write-Host "Creating Virtual Network: $($VNET.Item('Name')) in $ResourceGroupName" -ForegroundColor Green
				$status = New-AzureRmResourceGroupDeployment -Name ($Subscription.'Deployment Name' + "-VNET") -ResourceGroupName $ResourceGroupName `
						-Mode Incremental `
						-TemplateParameterObject $VNET `
						-TemplateFile ("$template" + "$SAS") `
						-Force
				if ($status.ProvisioningState -eq 'Succeeded')
				{
					Write-Host "Success: Creating Virtual Network: $($VNET.Item('Name')) in $ResourceGroupName" -ForegroundColor Green
				}
				else 
				{
					throw "Warning: Creating Virtual Network: $($VNET.Item('Name')) in $($ResourceGroupName) is not in a Succeeded state, please validate"
				}
			}
		}
	}
	catch 
	{
		throw $_;
	}
}
function ConvertTo-Hashtable
{
	param
	(
		[object]$PsObject,
		[string[]]$Exclusionlist
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		$HashTable = New-Object hashtable;
		$Keys = $PsObject |Get-Member -MemberType NoteProperty |Select-Object -Property Name;

		foreach ($Key in $Keys)
		{
			if ($Key.Name -notin $Exclusionlist)
			{
				$HashTable.Add($Key.Name,$PsObject.($Key.Name));
			}
		}
		return $HashTable;
	}
	catch
	{
		throw $_;
	}
}