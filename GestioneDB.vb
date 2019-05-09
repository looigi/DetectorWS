Module GestioneDB
    Public Function LeggeQuery(ByVal Conn As Object, ByVal Sql As String, ByVal Connessione As String) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn, Connessione)
        Dim Rec As Object = CreateObject("ADODB.Recordset")

        Try
            Rec.Open(Sql, Conn)
        Catch ex As Exception
            Rec = Nothing
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Rec
    End Function

    Public Function ControllaAperturaConnessione(ByRef Conn As Object, ByVal Connessione As String) As Boolean
        Dim Ritorno As Boolean = False

        If Conn Is Nothing Then
            Ritorno = True
            Conn = ApreDB(Connessione)
        End If

        Return Ritorno
    End Function

    Public Sub ChiudeDB(ByVal TipoApertura As Boolean, ByRef Conn As Object)
        If TipoApertura = True Then
            Conn.Close()
        End If
    End Sub

    Public Function LeggeImpostazioniDiBase(ByVal Perc As String) As String
        Dim Connessione As String = ""

        ' Impostazioni di base
        Dim ListaConnessioni As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

        If ListaConnessioni.Count <> 0 Then
            ' Get the collection elements. 
            For Each Connessioni As ConnectionStringSettings In ListaConnessioni
                Dim Nome As String = Connessioni.Name
                Dim Provider As String = Connessioni.ProviderName
                Dim connectionString As String = Connessioni.ConnectionString

				If Nome = "ConnectionStringLOCALE" Then
					Connessione = "Provider=" & Provider & ";" & connectionString
					Exit For
				End If
			Next
        End If

        Return Connessione
    End Function

    Public Function ApreDB(ByVal Connessione As String) As Object
        ' Routine che apre il DB e vede se ci sono errori
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.CommandTimeout = 0
        Catch
            Conn = Nothing
        End Try

        Return Conn
    End Function

    Public Function EsegueSql(ByVal Conn As Object, ByVal Sql As String, ByVal Connessione As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn, Connessione)
        Dim Ritorno As String = ""

        ' Routine che esegue una query sul db
        Try
            Conn.Execute(Sql)
        Catch
            Ritorno = "Errore: " & Err.Description
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function EsegueSqlSenzaTRY(ByVal Conn As Object, ByVal Sql As String, ByVal Connessione As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn, Connessione)
        Dim Ritorno As String = ""

        ' Routine che esegue una query sul db
        Conn.Execute(Sql)

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function RitornaMaxID(NomeTabella As String, NomeID As String, Conn As Object, Connessione As String, idUtente As String) As Integer
        Dim Rec As Object = CreateObject("ADODB.Recordset")
        Dim Sql As String = ""
        Dim Numerello As String

        Sql = "Select Max(" & NomeID & ")+1 From " & NomeTabella & " Where idUtente=" & idUtente
        Rec = LeggeQuery(Conn, Sql, Connessione)
        If Rec(0).Value Is DBNull.Value = True Then
            Numerello = 1
        Else
            Numerello = Rec(0).Value
        End If
        Rec.Close()

        Return Numerello
    End Function
End Module
