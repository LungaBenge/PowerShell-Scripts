Unblock-File -Path "C:\Users\E948996\Documents\My Portfolio\Powershell\SQLExtract export to Spreadsheet.ps1";
Get-ExecutionPolicy RemoteSigned;
## ---------- Working with SQL Server ---------- ##

## - Get SQL Server Table data:
$SQLServer = 'srv003376';
$Database = 'PROD_Lamda_COPY_Enquiry';
$SqlQuery = @' '@;
 
## - Connect to SQL Server using non-SMO class 'System.Data':
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
$SqlConnection.ConnectionString = `
"Server = $SQLServer; Database = $Database; Integrated Security = True";
 
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
$SqlCmd.CommandText = $SqlQuery;
$SqlCmd.Connection = $SqlConnection;
 
## - Extract and build the SQL data object '$DataSetTable':
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
$SqlAdapter.SelectCommand = $SqlCmd;
$DataSet = New-Object System.Data.DataSet;
$SqlAdapter.Fill($DataSet);
$DataSetTable = $DataSet.Tables["Table"];
$DataSetTable

## ---------- Working with Excel ---------- ##
 
## - Create an Excel Application instance:
$xlsObj = New-Object -ComObject Excel.Application;
 
## - Create new Workbook and Sheet (Visible = 1 / 0 not visible)
$xlsObj.Visible = 0;
$xlsWb = $xlsobj.Workbooks.Add();
$xlsSh = $xlsWb.Worksheets.item(1);
 
## - Build the Excel column heading:
[Array] $getColumnNames = $DataSetTable.Columns | Select ColumnName;
 
## - Build column header:
[Int] $RowHeader = 1;
foreach ($ColH in $getColumnNames)
{
$xlsSh.Cells.item(1, $RowHeader).font.bold = $true;
$xlsSh.Cells.item(1, $RowHeader) = $ColH.ColumnName;
$RowHeader++;
};
 
## - Adding the data start in row 2 column 1:
[Int] $rowData = 2;
[Int] $colData = 1;
 
foreach ($rec in $DataSetTable.Rows)
{
foreach ($Coln in $getColumnNames)
{
## - Next line convert cell to be text only:
$xlsSh.Cells.NumberFormat = "@";
 
## - Populating columns:
$xlsSh.Cells.Item($rowData, $colData) = `
$rec.$($Coln.ColumnName).ToString();
$ColData++;
};
$rowData++; $ColData = 1;
};
 
## - Adjusting columns in the Excel sheet:
$xlsRng = $xlsSH.usedRange;
$xlsRng.EntireColumn.AutoFit();

## ---------- Saving file and Terminating Excel Application ---------- ##
 
## - Saving Excel file - if the file exist do delete then save
$xlsFile = `
"C:\Temp\NewExceldbResults_$((Get-Date).ToString("yyyyMMdd")).xls";
 
if (Test-Path $xlsFile)
{
Remove-Item $xlsFile
$xlsObj.ActiveWorkbook.SaveAs($xlsFile);
}
else
{
$xlsObj.ActiveWorkbook.SaveAs($xlsFile);
};
 
## Quit Excel and Terminate Excel Application process:
$xlsObj.Quit(); (Get-Process Excel*) | foreach ($_) { $_.kill() };
 
## - End of Script - ##