
// �ݒ�

var dsn         = "test_4d";
var db_username = "SQL_user";
var db_password = "sql";

// DSN�𗘗p����4D�ɐڑ�

var connection_string = "ODBC"
	+ ";Driver={4D v15 ODBC Driver 64-bit}"
	+ ";DSN=" + dsn
	+ ";UID=" + db_username
	+ ";PWD=" + db_password
	+ ";"
;
var con = WScript.CreateObject("ADODB.Connection");
con.Open( connection_string );

// �g�����U�N�V�������J�n���ă��R�[�h���쐬

con.BeginTrans();

var cmd = new ActiveXObject("ADODB.Command");
cmd.ActiveConnection = con;
cmd.CommandType = 1;
cmd.Prepared = true;

cmd.CommandText = "INSERT INTO Event_Log (Call_event) VALUES ('test');";
cmd.Execute();

con.CommitTrans();

// �㏈��
con.Close();