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

				$TemplateParameterObject = ConvertTo-Hashtable -PsObject $OmsWorkspaceData -Exclusionlist @('OMSRSG','Template','SAS');
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