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