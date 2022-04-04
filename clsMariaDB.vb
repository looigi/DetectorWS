



Imports MySqlConnector

Public Class clsMariaDB
	'Dim Conn As MySqlConnection
	Dim StringaConnessione As String

	Public Function apreConnessione(c) As Object
		'StringaConnessione = c
		'Conn = New MySqlConnection(c)

		'Try
		'	Conn.Open()
		'Catch ex As Exception
		'	Return ex.Message
		'End Try

		'Return Conn
	End Function

	Public Function ChiudiConn(Conn)
		Conn.Close()

		Return "OK"
	End Function

	Public Function EsegueSql(Sql As String, ModificaQuery As Boolean) As String
		'Dim Errore As String = ""

		'' Routine che esegue una query sul db
		'Try
		'	Dim cmd As MySqlCommand = New MySqlCommand(Sql, Conn)

		'	'If Sql.ToUpper.Contains("Insert Into Arbitri ".ToUpper) Then
		'	'	Errore = Sql
		'	'	Return Errore
		'	'End If
		'	cmd.ExecuteNonQuery()

		'	Errore = "OK"
		'Catch ex As MySqlException
		'	Errore = ex.Message
		'End Try

		'Return Errore
	End Function

	Public Function Lettura(sql As String, ModificaQuery As Boolean) As Object
		'Dim cmd As MySqlCommand = New MySqlCommand(sql, Conn)
		'Dim Ritorno As MySqlDataReader
		'Dim rec As Object = Nothing

		'Try
		'	Ritorno = cmd.ExecuteReader()
		'	Dim ds As DataSet = New DataSet()
		'	Dim theCommand As New DataTable()
		'	ds.Tables.Add(theCommand)
		'	ds.EnforceConstraints = False
		'	theCommand.Load(Ritorno)
		'	rec = New clsRecordset(theCommand, sql)
		'Catch ex As Exception
		'	rec = "MDB ERROR:" & ex.Message
		'End Try

		'Return rec
	End Function

End Class
