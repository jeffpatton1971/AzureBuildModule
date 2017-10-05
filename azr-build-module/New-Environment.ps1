param
(
	[string]$Path,
	[System.Management.Automation.PSCredential]$Credential
)
$SheetArray = @('Subscriptions','Environments','ResourceGroups','VirtualNetworks','OMSWOrkspaces','WebApps','TrafficManager','StorageAccounts')

Import-Module .\BuildSheet\BuildSheet.psd1;
Import-Module .\BuildDevice\BuildDevice.psd1;

$SubscriptionData = Get-BuildSheetData -Path $Path -Worksheet 'Subscriptions';

Login-AzureRmAccount -Credential $Credential -SubscriptionId $SubscriptionData.'Subscription ID' -TenantId $SubscriptionData.'Tenant ID'

foreach ($sheet in $SheetArray)
{
	switch ($Sheet)
	{
		'Environments'
		{
			$EnvironmentData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'ResourceGroups'
		{
			$ResourceGroupData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
			New-ResourceGroup -ResourceGroupData $ResourceGroupData -SubscriptionData $SubscriptionData;
		}
		'VirtualNetworks'
		{
			$VirtualNetworkData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'OMSWorkspaces'
		{
			$OmsWorkspaceData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
			New-OmsWorkspace -OmsWorkspaceData $OmsWorkspaceData -SubscriptionData $SubscriptionData;
		}
		'WebApps'
		{
			$WebAppData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
			New-WebApp -WebAppData $WebAppData -SubscriptionData $SubscriptionData;
		}
		'TrafficManager'
		{
			$TrafficManagerData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'StorageAccounts'
		{
			$StorageAccountData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
	}
}