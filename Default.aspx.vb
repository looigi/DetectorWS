Imports System.IO
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports Ionic.Zip
Imports System.Threading

Public Class _Default
	Inherits System.Web.UI.Page

	Public Structure dati
		Dim gf As GestioneFilesDirectory
		Dim myFileCollection As HttpFileCollection
		Dim vNomeFile As String
		Dim FileLog As String
		Dim NomeFileLog As String
	End Structure

	Private Sub form1_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Init
		Dim vNomeFile As String = ""
		Dim gf As New GestioneFilesDirectory
		Dim FileLog As String = Server.MapPath(".") & "\Scaricati\Log"
		Dim NomeFileLog As String = "Log" & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".txt"

		gf.CreaDirectoryDaPercorso(FileLog & "\")
		' gf.ApreFileDiTestoPerScrittura(FileLog & "\" & NomeFileLog)
		' gf.ScriveTestoSuFileAperto("-----------------------------------------------")
		ScriveLog(FileLog, NomeFileLog, "-----------------------------------------------")
		ScriveLog(FileLog, NomeFileLog, "Arrivata nuova richiesta")

		If Not String.IsNullOrEmpty(Request.Form("nomefile")) Then
			vNomeFile = Request.Form("nomefile")
		End If

		If vNomeFile <> "" Then
			ScriveLog(FileLog, NomeFileLog, "Nome file: " & vNomeFile)

			Dim MyFileCollection As HttpFileCollection = Request.Files

			If MyFileCollection.Count > 0 Then
				Dim STthread_Data As dati
				STthread_Data.gf = gf
				STthread_Data.myFileCollection = MyFileCollection
				STthread_Data.NomeFileLog = NomeFileLog
				STthread_Data.FileLog = FileLog
				STthread_Data.vNomeFile = vNomeFile

				ScriveLog(FileLog, NomeFileLog, "Arrivato zip: " & vNomeFile)

				ElaboraInizio(STthread_Data)

				Dim multiThread As Thread = New Thread(AddressOf Elabora)
				multiThread.SetApartmentState(ApartmentState.MTA)
				multiThread.Start(STthread_Data)

				' Elabora(gf, MyFileCollection, vNomeFile, FileLog, NomeFileLog)
			Else
				ScriveLog(FileLog, NomeFileLog, "Nessuno streaming file arrivato")
			End If
		Else
			ScriveLog(FileLog, NomeFileLog, "Nessun nome file arrivato")
		End If

		' gf.ChiudeFileDiTestoDopoScrittura()
		gf = Nothing
	End Sub

	Private Sub ScriveLog(FileLog As String, NomeFileLog As String, Cosa As String)
		Dim gf As New GestioneFilesDirectory
		gf.ApreFileDiTestoPerScrittura(FileLog & "\" & NomeFileLog)
		gf.ScriveTestoSuFileAperto(Now & ": " & Cosa)
		gf.ChiudeFileDiTestoDopoScrittura()
	End Sub

	Private Sub ElaboraInizio(strDati As dati)
		Dim gf As GestioneFilesDirectory = strDati.gf
		Dim vNomeFile As String = strDati.vNomeFile
		Dim FileLog As String = strDati.FileLog
		Dim NomeFileLog As String = strDati.NomeFileLog
		Dim MyFileCollection As HttpFileCollection = strDati.myFileCollection

		ScriveLog(FileLog, NomeFileLog, "Entrato in ElaboraInizio")

		Dim FilePath As String = Server.MapPath(".") & "\Scaricati"
		Dim FileUnzip As String = Server.MapPath(".") & "\Scaricati\Unzip"
		Dim FileArchiviati As String = Server.MapPath(".") & "\Scaricati\Archiviati"
		Dim FileRovinati As String = Server.MapPath(".") & "\Scaricati\Rovinati"
		Dim FileBackup As String = Server.MapPath(".") & "\Scaricati\Backup"
		Dim Contatore As Integer = 0
		Dim Altro As String = ""

		gf.CreaDirectoryDaPercorso(FilePath & "\")
		gf.CreaDirectoryDaPercorso(FileArchiviati & "\")
		gf.CreaDirectoryDaPercorso(FileUnzip & "\")
		gf.CreaDirectoryDaPercorso(FileRovinati & "\")
		gf.CreaDirectoryDaPercorso(FileBackup & "\")

		If Not MyFileCollection Is Nothing Then
			If File.Exists(FilePath & "\" & vNomeFile) Then
				ScriveLog(FileLog, NomeFileLog, "Elimino il file già esistente in " & FilePath & "\" & vNomeFile)
				File.Delete(FilePath & "\" & vNomeFile)
			End If

			ScriveLog(FileLog, NomeFileLog, "Salvo il file in " & FilePath & "\" & vNomeFile)
			MyFileCollection(0).SaveAs(FilePath & "\" & vNomeFile)

			Dim est As String = gf.TornaEstensioneFileDaPath(vNomeFile)
			vNomeFile = vNomeFile.Replace(est, "")
			vNomeFile = vNomeFile & "_" & Now.Year & Format(Now.Month, 0) & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
			vNomeFile &= est

			ScriveLog(FileLog, NomeFileLog, "Salvo il file in " & FileBackup & "\" & vNomeFile)
			MyFileCollection(0).SaveAs(FileBackup & "\" & vNomeFile)
		End If
	End Sub

	Private Sub Elabora(strDati As dati)
		Dim gf As GestioneFilesDirectory = strDati.gf
		Dim FileLog As String = strDati.FileLog
		Dim NomeFileLog As String = strDati.NomeFileLog
		Dim FilePath As String = Server.MapPath(".") & "\Scaricati"
		Dim vNomeFile As String = strDati.vNomeFile
		Dim FileUnzip As String = Server.MapPath(".") & "\Scaricati\Unzip"
		Dim FileArchiviati As String = Server.MapPath(".") & "\Scaricati\Archiviati"
		Dim FileRovinati As String = Server.MapPath(".") & "\Scaricati\Rovinati"

		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))
		ScriveLog(FileLog, NomeFileLog, "Lettura connessione: " & Connessione)

		If Connessione = "" Then
			ScriveLog(FileLog, NomeFileLog, "Problemi nella lettura della connessione")
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If Conn Is Nothing Then
				ScriveLog(FileLog, NomeFileLog, "Problemi nell'apertura del db")
			Else
				ScriveLog(FileLog, NomeFileLog, "Db aperto correttamente")

				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""

				If File.Exists(FilePath & "\" & vNomeFile) Then
					ScriveLog(FileLog, NomeFileLog, "Elaboro il file arrivato")

					Dim Quanti As Integer = 0

					ScriveLog(FileLog, NomeFileLog, "Elimino i files già presenti nella cartella di unzip")
					gf.ScansionaDirectorySingola(FileUnzip)

					Dim qq As Integer = gf.RitornaQuantiFilesRilevati
					Dim ff() As String = gf.RitornaFilesRilevati

					For i As Integer = 1 To qq
						gf.EliminaFileFisico(ff(i))
					Next

					Try
						Using zip As ZipFile = ZipFile.Read(FilePath & "\" & vNomeFile)
							For Each zip_entry As ZipEntry In zip
								ScriveLog(FileLog, NomeFileLog, "Unzip " & zip_entry.FileName)
								zip_entry.Extract(FileUnzip)
								Quanti += 1
							Next zip_entry
						End Using
					Catch ex As Exception
						ScriveLog(FileLog, NomeFileLog, "Unzip error: " & ex.Message)
					End Try

					If Quanti > 0 Then
						ScriveLog(FileLog, NomeFileLog, "Files unzippati: " & Quanti)

						gf.LeggeFilesDaDirectory(FileUnzip)

						Dim q As Integer = gf.RitornaQuantiFilesRilevati
						Dim filetti() As String = gf.RitornaFilesRilevati

						For i As Integer = 1 To q
							Dim Tipo As String

							If filetti(i).Contains("LL_") Then
								ScriveLog(FileLog, NomeFileLog, "Elaborazione file posizioni " & i & "/" & q & ": " & filetti(i))
								Tipo = "LL"
							Else
								ScriveLog(FileLog, NomeFileLog, "Elaborazione file multimedia " & i & "/" & q & ": " & filetti(i))
								Tipo = "MM"
							End If

							Dim contenuto As String = gf.LeggeFileIntero(filetti(i))

							If contenuto = "" Then
								ScriveLog(FileLog, NomeFileLog, "         File vuoto. Skippo.")
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
										ScriveLog(FileLog, NomeFileLog, "Eliminazione righe posizioni per giorno " & Datella)
										Sql = "Delete From Posizioni Where DataPos = '" & Datella & "'"
										EsegueSql(Conn, Sql, Connessione)

										ScriveLog(FileLog, NomeFileLog, "Inserimento righe posizioni per giorno " & Datella)
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
												'ScriveLog(FileLog, NomeFileLog, "SQL:" & Sql)
												Dim Ritorno As String = EsegueSql(Conn, Sql, Connessione)
												'ScriveLog(FileLog, NomeFileLog, "Ritorno:" & Ritorno)

												If Ritorno.Contains("Errore: ") Then
													Sql = "Delete From Posizioni Where DataPos = '" & Datella & "'"
													'ScriveLog(FileLog, NomeFileLog, "SQL:" & Sql)
													EsegueSql(Conn, Sql, Connessione)
													Ok = False

													Exit For
												End If

												Quante += 1
											End If
										Next
										ScriveLog(FileLog, NomeFileLog, "Righe posizioni inserite per giorno: " & Quante)
									Case "MM"
										' Scrive i dati su sql server per i multimedia
										ScriveLog(FileLog, NomeFileLog, "Eliminazione righe multimedia per giorno " & Datella)
										Sql = "Delete From Multimedia Where DataPos = '" & Datella & "'"
										'ScriveLog(FileLog, NomeFileLog, "SQL:" & Sql)
										EsegueSql(Conn, Sql, Connessione)

										ScriveLog(FileLog, NomeFileLog, "Inserimento righe multimedia per giorno " & Datella)
										For Each Riga As String In Righe
											If Riga.Trim <> "" Then
												Dim Campi() As String = Riga.Split(";")
												If Riga.Contains("V") Then
													Dim Lat As String = Campi(0)
													Dim Lon As String = Campi(1)
													Dim Quando As String = Campi(2)
													Dim NomeFile As String = Campi(3)
													Dim Tipologia As String = Campi(4)

													Sql = "Insert Into Multimedia Values (" &
														"'" & Datella & "', " &
														"'" & Quando & "', " &
														"'" & Lat & "', " &
														"'" & Lon & "', " &
														"'" & NomeFile & "'," &
														"'" & Tipologia & "'" &
														")"
												Else
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
												End If
												'ScriveLog(FileLog, NomeFileLog, "SQL:" & Sql)
												Dim Ritorno As String = EsegueSql(Conn, Sql, Connessione)

												If Ritorno.Contains("Errore: ") Then
													Sql = "Delete From Multimedia Where DataPos = '" & Datella & "'"
													'ScriveLog(FileLog, NomeFileLog, "SQL:" & Sql)
													EsegueSql(Conn, Sql, Connessione)
													Ok = False

													Exit For
												End If

												Quante += 1
											End If
										Next
										ScriveLog(FileLog, NomeFileLog, "Righe multimedia inserite per giorno: " & Quante)
								End Select

								If Ok Then
									ScriveLog(FileLog, NomeFileLog, "Elaborazione effettuata. Sposto il file in archivio")
									Dim dest As String = filetti(i)
									dest = gf.TornaNomeFileDaPath(dest)
									dest = FileArchiviati & "\" & dest

									gf.CopiaFileFisico(filetti(i), dest, True)
									If File.Exists(dest) Then
										ScriveLog(FileLog, NomeFileLog, "Spostato il file in archivio. Lo elimino dall'origine")
										gf.EliminaFileFisico(filetti(i))
									Else
										ScriveLog(FileLog, NomeFileLog, "NON Spostato il file in archivio")
									End If
								Else
									ScriveLog(FileLog, NomeFileLog, "Elaborazione NON effettuata. Sposto il file in non caricati")
									Dim dest As String = filetti(i)
									dest = gf.TornaNomeFileDaPath(dest)
									dest = FileRovinati & "\" & dest

									gf.CopiaFileFisico(filetti(i), dest, True)
									If File.Exists(dest) Then
										ScriveLog(FileLog, NomeFileLog, "Spostato il file in non caricati. Lo elimino dall'origine")
										gf.EliminaFileFisico(filetti(i))
									Else
										ScriveLog(FileLog, NomeFileLog, "NON Spostato il file in archivio")
									End If
								End If
							End If
						Next

						ScriveLog(FileLog, NomeFileLog, "Elaborazione terminata. Elimino il file arrivato")
						gf.EliminaFileFisico(FilePath & "\" & vNomeFile)
					Else
						ScriveLog(FileLog, NomeFileLog, "Nessun file unzippato")
					End If
				Else
					ScriveLog(FileLog, NomeFileLog, "Nessun file salvato")
				End If
			End If
		End If
	End Sub

	Protected Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
		Dim gf As New GestioneFilesDirectory
		Dim FileLog As String = Server.MapPath(".") & "\Scaricati\Log"
		Dim NomeFileLog As String = "Log" & Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".txt"

		' gf.ApreFileDiTestoPerScrittura(FileLog & "\" & NomeFileLog)
		' gf.ScriveTestoSuFileAperto("-----------------------------------------------")
		ScriveLog(FileLog, NomeFileLog, "Creato nuovo test")

		' Me.Elabora(gf, Nothing, "Daunload.zip", FileLog, NomeFileLog)
		Dim STthread_Data As dati
		STthread_Data.gf = gf
		STthread_Data.myFileCollection = Nothing
		STthread_Data.NomeFileLog = NomeFileLog
		STthread_Data.FileLog = FileLog
		STthread_Data.vNomeFile = "Daunload.zip"

		Dim multiThread As Thread = New Thread(AddressOf Elabora)
		multiThread.SetApartmentState(ApartmentState.MTA)
		multiThread.Start(STthread_Data)

		'gf.ChiudeFileDiTestoDopoScrittura()
		gf = Nothing
	End Sub
End Class