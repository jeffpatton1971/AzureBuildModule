# Buildsheet Module
```
Import-Module .\Rackspace\racker\azr-build-module\azr-build-module\BuildSheet\BuildSheet\BuildSheet.psd1 -Force

Get-Command -Module BuildSheet

CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Get-BuildSheetData                                 1.0        BuildSheet

NAME
    Get-BuildSheetData

SYNOPSIS
    A function to returned data from sheets in an Excel spreadsheet


SYNTAX
    Get-BuildSheetData [-Path] <String> [-Worksheet] <String> [-Range] <String> [<CommonParameters>]


DESCRIPTION
    This function leverages OLEDB to open an Excel spreadsheet as a datasource
    and query it the same way you would a database. This method of access
    improves performance, and provides less complexity.


PARAMETERS
    -Path <String>
        The path and filename of the buildsheet to process

        Required?                    true
        Position?                    2
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -Worksheet <String>
        This is a defined list of worksheets

        Required?                    true
        Position?                    3
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -Range <String>
        This is a regular Excel Range notation (A1:B12)

        Required?                    true
        Position?                    4
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

INPUTS

OUTPUTS

NOTES


        You may need to download the x64 ACE provider, see the first link

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>Get-BuildSheetData -Path 'D:\Documents\Build Sheet-170925-08536.xlsx' -Worksheet Environments

    Environment       Location         SubscriptionName
    -----------       --------         ----------------
    Production        South Central US Solarwinds Azure - Aviator Support
    Q/A               South Central US Solarwinds Azure - Aviator Support
    Development       South Central US Solarwinds Azure - Aviator Support
    Disaster Recovery East US          Solarwinds Azure - Aviator Support

    Description
    ===========
    This example shows how to pull the data from the Environments tab





RELATED LINKS
    https://www.microsoft.com/en-us/download/details.aspx?id=13255
    https://msdn.microsoft.com/en-us/library/system.data.oledb(v=vs.110).aspx
    https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbconnection(v=vs.110).aspx
    https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbcommand(v=vs.110).aspx
    https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbdatareader(v=vs.110).aspx
```    
# Troubshooting

If you see the message below you will need to download the ACE OLDB driver

https://www.microsoft.com/en-us/download/details.aspx?id=13255

```
Exception calling "Open" with "0" argument(s): "The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine."
At C:\Rackspace\azr-build-module\azr-build-module\BuildSheet\BuildSheet.psm1:65 char:3
+         $OleDbConnection.Open();
+         ~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
    + FullyQualifiedErrorId : InvalidOperationException
```
