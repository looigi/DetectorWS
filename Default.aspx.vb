Imports System.IO
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports Ionic.Zip

Public Class _Default
	Inherits System.Web.UI.Page

	Private Sub form1_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Init
		Dim vNomeFile As String = ""
		Dim gf As New GestioneFilesDirectory
		Dim FileLog As String = Server.MapPath(".") & "\Scaricati\Log"
		Dim NomeFileLog As String = "Log.txt"

		gf.CreaDirectoryDaPercorso(FileLog & "\")
		gf.ApreFileDiTestoPerScrittura(FileLog & "\" & NomeFileLog)
		gf.ScriveTestoSuFileAperto("-----------------------------------------------")
		ScriveLog(gf, "Arrivata nuova richiesta")

		If Not String.IsNullOrEmpty(Request.Form("nomefile")) Then
			vNomeFile = Request.Form("nomefile")
		End If

		If vNomeFile <> "" Then
			ScriveLog(gf, "Nome file: " & vNomeFile)

			Dim MyFileCollection As HttpFileCollection = Request.Files

			If MyFileCollection.Count > 0 Then
				Elabora(gf, MyFileCollection, vNomeFile)
			Else
				ScriveLog(gf, "Nessuno streaming file arrivato")
			End If
		Else
			ScriveLog(gf, "Nessun nome file arrivato")
		End If

		gf.ChiudeFileDiTestoDopoScrittura()
		gf = Nothing
	End Sub

	Private Sub ScriveLog(gf As GestioneFilesDirectory, Cosa As String)
		gf.ScriveTestoSuFileAperto(Now & ": " & Cosa)
	End Sub

	Private Sub Elabora(gf As GestioneFilesDirectory, MyFileCollection As HttpFileCollection, vNomeFile As String)
		Dim FilePath As String = Server.MapPath(".") & "\Scaricati"
		Dim FileUnzip As String = Server.MapPath(".") & "\Scaricati\Unzip"
		Dim FileArchiviati As String = Server.MapPath(".") & "\Scaricati\Archiviati"
		Dim FileRovinati As String = Server.MapPath(".") & "\Scaricati\Rovinati"
		Dim Contatore As Integer = 0
		Dim Altro As String = ""

		gf.CreaDirectoryDaPercorso(FilePath & "\")
		gf.CreaDirectoryDaPercorso(FileArchiviati & "\")
		gf.CreaDirectoryDaPercorso(FileUnzip & "\")
		gf.CreaDirectoryDaPercorso(FileRovinati & "\")

		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

		If Connessione = "" Then
			ScriveLog(gf, "Problemi nella lettura della connessione")
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If Conn Is Nothing Then
				ScriveLog(gf, "Problemi nell'apertura del db")
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				If Not MyFileCollection Is Nothing Then
					If File.Exists(FilePath & "\" & vNomeFile) Then
						ScriveLog(gf, "Elimino il file già esistente in " & FilePath & "\" & vNomeFile)
						File.Delete(FilePath & "\" & vNomeFile)
					End If

					ScriveLog(gf, "Salvo il file in " & FilePath & "\" & vNomeFile)
					MyFileCollection(0).SaveAs(FilePath & "\" & vNomeFile)
				End If

				If File.Exists(FilePath & "\" & vNomeFile) Then
					ScriveLog(gf, "Elaboro il file arrivato")

					Dim Quanti As Integer = 0

					ScriveLog(gf, "Elimino i files già presenti nella cartella di unzip")
					gf.ScansionaDirectorySingola(FileUnzip)

					Dim qq As Integer = gf.RitornaQuantiFilesRilevati
					Dim ff() As String = gf.RitornaFilesRilevati

					For i As Integer = 1 To qq
						gf.EliminaFileFisico(ff(i))
					Next

					Try
						Using zip As ZipFile = ZipFile.Read(FilePath & "\" & vNomeFile)
							For Each zip_entry As ZipEntry In zip
								ScriveLog(gf, "Unzip " & zip_entry.FileName)
								zip_entry.Extract(FileUnzip)
								Quanti += 1
							Next zip_entry
						End Using
					Catch ex As Exception
						ScriveLog(gf, "Unzip error: " & ex.Message)
					End Try

					If Quanti > 0 Then
						ScriveLog(gf, "Files unzippati: " & Quanti)

						gf.LeggeFilesDaDirectory(FileUnzip)

						Dim q As Integer = gf.RitornaQuantiFilesRilevati
						Dim filetti() As String = gf.RitornaFilesRilevati

						For i As Integer = 1 To q
							Dim Tipo As String

							If filetti(i).Contains("LL_") Then
								ScriveLog(gf, "Elaborazione file posizioni " & i & "/" & q & ": " & filetti(i))
								Tipo = "LL"
							Else
								ScriveLog(gf, "Elaborazione file multimedia " & i & "/" & q & ": " & filetti(i))
								Tipo = "MM"
							End If

							Dim contenuto As String = gf.LeggeFileIntero(filetti(i))

							If contenuto = "" Then
								ScriveLog(gf, "         File vuoto. Skippo.")
							Else
								Dim Ok As Boolean = True
								Dim Righe() As String = contenuto.Split("§")

								Dim sData As String = filetti(i)
								sData = gf.TornaNomeFileDaPath(sData)
								sData = sData.Replace(Tipo & "_", "")
								sData = sData.Replace(".txt", "")
								Dim dd() As String = sData.Split("_")
								Dim Datella As String = dd(2) & "-" & dd(1) & "-" & dd(0)
								Dim Quante As Integer = 0

								Select Case Tipo
									Case "LL"
										' Scrive i dati su sql server per le posizioni
										ScriveLog(gf, "Eliminazione righe posizioni per giorno " & Datella)
										Sql = "Delete From Posizioni Where DataPos = '" & Datella & "'"
										EsegueSql(Conn, Sql, Connessione)

										ScriveLog(gf, "Inserimento righe posizioni per giorno " & Datella)
										For Each Riga As String In Righe
											If Riga.Trim <> "" Then
												Dim Campi() As String = Riga.Split(";")
												Dim Lat As String = Campi(0)
												Dim Lon As String = Campi(1)
												Dim Quando As String = Campi(2)
												Dim Velocita As String = Campi(3)

												Sql = "Insert Into Posizioni Values (" &
													"'" & Datella & "', " &
													"'" & Quando & "', " &
													"'" & Lat & "', " &
													"'" & Lon & "', " &
													"'" & Velocita & "'" &
													")"
												Dim Ritorno As String = EsegueSql(Conn, Sql, Connessione)

												If Ritorno.Contains("Errore: ") Then
													Sql = "Delete From Posizioni Where DataPos = '" & Datella & "'"
													EsegueSql(Conn, Sql, Connessione)
													Ok = False

													Exit For
												End If

												Quante += 1
											End If
										Next
										ScriveLog(gf, "Righe posizioni inserite per giorno: " & Quante)
									Case "MM"
										' Scrive i dati su sql server per i multimedia
										ScriveLog(gf, "Eliminazione righe multimedia per giorno " & Datella)
										Sql = "Delete From Multimedia Where DataPos = '" & Datella & "'"
										EsegueSql(Conn, Sql, Connessione)

										ScriveLog(gf, "Inserimento righe multimedia per giorno " & Datella)
										For Each Riga As String In Righe
											If Riga.Trim <> "" Then
												Dim Campi() As String = Riga.Split(";")
												Dim Lat As String = Campi(0)
												Dim Lon As String = Campi(1)
												Dim Quando As String = Campi(2)
												Dim NomeFile As String = Campi(5)
												Dim Tipologia As String = Campi(6)

												Sql = "Insert Into Multimedia Values (" &
													"'" & Datella & "', " &
													"'" & Quando & "', " &
													"'" & Lat & "', " &
													"'" & Lon & "', " &
													"'" & NomeFile & "'," &
													"'" & Tipologia & "'" &
													")"
												Dim Ritorno As String = EsegueSql(Conn, Sql, Connessione)

												If Ritorno.Contains("Errore: ") Then
													Sql = "Delete From Multimedia Where DataPos = '" & Datella & "'"
													EsegueSql(Conn, Sql, Connessione)
													Ok = False

													Exit For
												End If

												Quante += 1
											End If
										Next
										ScriveLog(gf, "Righe multimedia inserite per giorno: " & Quante)
								End Select

								If Ok Then
									ScriveLog(gf, "Elaborazione effettuata. Sposto il file in archivio")
									Dim dest As String = filetti(i)
									dest = gf.TornaNomeFileDaPath(dest)
									dest = FileArchiviati & "\" & dest

									gf.CopiaFileFisico(filetti(i), dest, True)
									If File.Exists(dest) Then
										ScriveLog(gf, "Spostato il file in archivio. Lo elimino dall'origine")
										gf.EliminaFileFisico(filetti(i))
									Else
										ScriveLog(gf, "NON Spostato il file in archivio")
									End If
								Else
									ScriveLog(gf, "Elaborazione NON effettuata. Sposto il file in non caricati")
									Dim dest As String = filetti(i)
									dest = gf.TornaNomeFileDaPath(dest)
									dest = FileRovinati & "\" & dest

									gf.CopiaFileFisico(filetti(i), dest, True)
									If File.Exists(dest) Then
										ScriveLog(gf, "Spostato il file in non caricati. Lo elimino dall'origine")
										gf.EliminaFileFisico(filetti(i))
									Else
										ScriveLog(gf, "NON Spostato il file in archivio")
									End If
								End If
							End If
						Next

						ScriveLog(gf, "Elaborazione terminata. Elimino il file arrivato")
						gf.EliminaFileFisico(FilePath & "\" & vNomeFile)
					Else
						ScriveLog(gf, "Nessun file unzippato")
					End If
				Else
					ScriveLog(gf, "Nessun file salvato")
				End If
			End If
		End If
	End Sub

	Protected Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
		Dim gf As New GestioneFilesDirectory
		Dim FileLog As String = Server.MapPath(".") & "\Scaricati\Log"
		Dim NomeFileLog As String = "Log.txt"

		gf.ApreFileDiTestoPerScrittura(FileLog & "\" & NomeFileLog)
		gf.ScriveTestoSuFileAperto("-----------------------------------------------")
		ScriveLog(gf, "Creato nuovo test")

		Elabora(gf, Nothing, "Daunload.zip")

		gf.ChiudeFileDiTestoDopoScrittura()
		gf = Nothing
	End Sub
End Class