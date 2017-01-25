
// 設定

var dsn         = "test_4d";
var db_username = "SQL_user";
var db_password = "sql";

// DSNを利用して4Dに接続

var connection_string = "ODBC"
	+ ";Driver={4D v15 ODBC Driver 64-bit}"
	+ ";DSN=" + dsn
	+ ";UID=" + db_username
	+ ";PWD=" + db_password
	+ ";"
;
var con = WScript.CreateObject("ADODB.Connection");
con.Open( connection_string );

// トランザクションを開始してレコードを作成

con.BeginTrans();

var cmd = new ActiveXObject("ADODB.Command");
cmd.ActiveConnection = con;
cmd.CommandType = 1;
cmd.Prepared = true;

cmd.CommandText = "INSERT INTO Event_Log (Call_event) VALUES ('test');";
cmd.Execute();

con.CommitTrans();

// 後処理
con.Close();