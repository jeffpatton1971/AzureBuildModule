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
		.PARAMETER Range
			This is a regular Excel Range notation (A1:B12)
		.EXAMPLE

      Get-BuildSheetData -Path D:\Documents\170926-07791.xlsx -Worksheet Environments -Range A1:C3

      Environment       Location SubscriptionName
      -----------       -------- ----------------
      Production        East US  Paradigm Software - Aviator
      Disaster Recovery West US  Paradigm Software - Aviator

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
    [string]$Worksheet,
    [Parameter(Mandatory=$True,Position=3)]
    [string]$Range
	)
	try
	{
		$ErrorActionPreference = 'Stop';
		$Error.Clear();

		$Provider = "Provider=Microsoft.ACE.OLEDB.12.0";
		$DataSource = "Data Source = $($Path)";
		$Properties = "Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`"";
		$OleDbConnection = New-Object System.Data.OleDb.OleDbConnection("$Provider;$DataSource;$Properties");

    $Query = "SELECT * FROM [$($Worksheet)$" + $Range + "]";
    
    $DataSet = New-Object System.Data.DataSet;
    $oleDbDataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter($Query,$OleDbConnection.ConnectionString);
    $oleDbDataAdapter.Fill($DataSet) |Out-Null;
    $DataTable = New-Object System.Data.DataTable;
    $DataTable = $DataSet.Tables[0];

		$OleDbConnection.close();
		$OleDbConnection.Dispose();

		Return $DataTable;
	}
	catch
	{
		throw $_;
	}
}