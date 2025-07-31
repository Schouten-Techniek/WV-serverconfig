$Conn = new-object System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\inetpub\wwwroot\WEBVisionNT\PROJEKTE\Beheer\Trend\trendDef.mdb'");
$conn.open();
$da = new-object System.Data.OleDb.OleDbDataAdapter(new-object System.Data.OleDb.OleDbCommand("SELECT ALL FID,CID,DP_AKT,DP_ADR,DP_VARI,DP_SKAL,DP_TXT,DP_START,trendUser,trendDate FROM DP_DEF",$Conn));
$dt = new-object System.Data.dataTable;
$da.fill($dt) | Out-Null;
$output=$dt | ConvertTo-csv -Delimiter ';';
$conn.close();
$output;