$Conn = new-object System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\inetpub\wwwroot\WEBVisionNT\PROJEKTE\Beheer\Daten\PVII.mdb'");
$conn.open();
$da = new-object System.Data.OleDb.OleDbDataAdapter(new-object System.Data.OleDb.OleDbCommand("SELECT Liegenschaft, Plant, Kreis, DIF, C_Adr  FROM MainTab",$Conn));
$dt = new-object System.Data.dataTable;
$da.fill($dt) | Out-Null;
$output=$dt | ConvertTo-csv -Delimiter ';';
$conn.close();
$output;