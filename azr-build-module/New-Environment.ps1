param
(
	[string]$Path,
	[System.Management.Automation.PSCredential]$Credential
)
$SheetArray = @('Subscriptions','Environments','ResourceGroups','VirtualNetworks','OMSWOrkspaces','WebApps','TrafficManager','StorageAccounts')

Import-Module .\BuildSheet\BuildSheet.psd1;

foreach ($sheet in $SheetArray)
{
	switch ($Sheet)
	{
		'Subscriptions'
		{
			$SubscriptionData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'Environments'
		{
			$EnvironmentData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'ResourceGroups'
		{
			$ResourceGroupData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'VirtualNetworks'
		{
			$VirtualNetworkData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'OMSWorkspaces'
		{
			$OmsWorkspaceData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
		}
		'WebApps'
		{
			$WebAppData = Get-BuildSheetData -Path $Path -Worksheet $Sheet;
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