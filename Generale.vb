Module Generale
	Public TipoDB As String = "MARIADB"
	Public StringaErrore As String = "ERROR: "

	Public Function LeggeImpostazioniDiBase(Percorso As String) As String
		Dim Connessione As String = ""

		' Impostazioni di base
		Dim ListaConnessioni As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

		If ListaConnessioni.Count <> 0 Then
			' Get the collection elements. 
			For Each Connessioni As ConnectionStringSettings In ListaConnessioni
				Dim Nome As String = Connessioni.Name
				Dim Provider As String = Connessioni.ProviderName
				Dim connectionString As String = Connessioni.ConnectionString

				If TipoDB = "SQLSERVER" Then
					If Nome = "SQLConnectionStringLOCALESS" Then
						Connessione = "Provider=" & Provider & ";" & connectionString
						Connessione = Replace(Connessione, "*^*^*", Percorso & "\")
						Exit For
					End If
				Else
					If Nome = "SQLConnectionStringLOCALEMD" Then
						Connessione = connectionString
						Connessione = Replace(Connessione, "*^*^*", Percorso & "\")
						Exit For
					End If
				End If
			Next
		End If

		Return Connessione
	End Function

End Module
