Function Get-BuildSheetData
{
	<#
		.SYNOPSIS
			A function to returned data from sheets in an Excel spreadsheet
		.DESCRIPTION
			This function leverages OLEDB to open an Excel spreadsheet as a datasource
			and query it the same way you would a database. This method of access
			improves performance, and provides less complexity.
		.PARAMETER Path
			The path and filename of the buildsheet to process
		.PARAMETER Worksheet
			This is a defined list of worksheets
		.EXAMPLE

			Get-BuildSheetData -Path 'D:\Documents\Build Sheet-170925-08536.xlsx' -Worksheet Environments

			Environment       Location         SubscriptionName
			-----------       --------         ----------------
			Production        South Central US Solarwinds Azure - Aviator Support
			Q/A               South Central US Solarwinds Azure - Aviator Support
			Development       South Central US Solarwinds Azure - Aviator Support
			Disaster Recovery East US          Solarwinds Azure - Aviator Support

			Description
			===========
			This example shows how to pull the data from the Environments tab			
		.NOTES
			You may need to download the x64 ACE provider, see the first link
		.LINK
			https://www.microsoft.com/en-us/download/details.aspx?id=13255
		.LINK
			https://msdn.microsoft.com/en-us/library/system.data.oledb(v=vs.110).aspx
		.LINK
			https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbconnection(v=vs.110).aspx
		.LINK
			https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbcommand(v=vs.110).aspx
		.LINK
			https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbdatareader(v=vs.110).aspx
	#>
	param
	(
		[Parameter(Mandatory=$True,Position=1)]
		[string]$Path,
		[Parameter(Mandatory=$True,Position=2)]
		[ValidateSet('ActiveDirectory','Environments','Lists','NSGs','OMSWorkspaces','RecoveryServicesVault','ResourceGroups','StorageAccounts','Subscriptions','TrafficManager','VirtualGateway','VirtualMachines','VirtualNetworks','VPNConnections','WebApps')]
		[string]$Worksheet
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		if (!(($Worksheet.Substring($Worksheet.Length-1,1)) -eq '$'))
		{
			$SheetName = "$($Worksheet)`$"
		}
		$Provider = "Provider=Microsoft.ACE.OLEDB.12.0";
		$DataSource = "Data Source = $($Path)";
		$Properties = "Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`"";
		$Query = "SELECT * FROM [$($SheetName)]";
		$OleDbConnection = New-Object System.Data.OleDb.OleDbConnection("$Provider;$DataSource;$Properties");
		$OleDbCommand = New-Object System.Data.OleDb.OleDbCommand($Query);
		$OleDbCommand.Connection = $OleDbConnection;
		$OleDbConnection.Open();

		$Columns = $OleDbConnection.GetSchema('Columns') |Where-Object -Property Table_Name -Like $SheetName;
		$OleDbDataReader = $OleDbCommand.ExecuteReader();

		$Data = @();
	
		while ($OleDbDataReader.Read())
		{
			$Item = New-Object -TypeName psobject;
			$Columns |Select-Object -ExpandProperty Column_Name |ForEach-Object {Add-Member -InputObject $Item -Name $_ -Value $OleDbDataReader.Item($_) -MemberType NoteProperty};
			$Data += $Item;
		}

		$OleDbDataReader.Close();
		$OleDbCommand.Dispose();
		$OleDbConnection.close();
		$OleDbConnection.Dispose();

		Return $Data;
	}
	catch
	{
		throw $_;
	}
}