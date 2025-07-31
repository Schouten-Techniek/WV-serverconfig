<%@ LANGUAGE="VBScript" %>
<!--#include virtual="/webvisionnt/zsystem/includes/doctype.asp"-->
<%
' ###################################################################################################
' Sprachsteuerungs - Session holen und dementsprechende Sprachdatei
' einlesen und ausführen (Konstanten registrieren/deklarieren...).
On Error Resume Next

Dim projekt
Dim d1
d1 = InStrRev(Session("Projektpfad"),"\")
projekt = mid(Session("Projektpfad"),d1+1,Len(Session("Projektpfad")))
' ###################################################################################################
%>

<!--#include file="../lang/Alarm/texteAlarm.asp"-->
<!--#include file="stringFunktionen.asp"-->
<!--#include file="ADOVBS.inc"-->
<!--#include file="IASUtil.asp"-->
<!--#include file="WEB.inc"-->
<!--#include file="WEB2.inc"-->
<!--#include file="priorityDescription_include.asp"-->

<html>
<head>

<title><%=ben013%>:&nbsp;<%=ChangeUmlInHTML(UCase(projekt))%></title>
</head>
<%
	'--Zwangsabmeldecounter zurücksetzen durch aktion---
	Session("inactivetime") = 0
	'---------------------------------------------------
	On Error Resume Next


	' ------------------------------------- [AUTOFILTER] -------------------------------------
	' Boolean um festzustellen ob Seite mit Autofilter geladen wird
	Dim autoFilter
	autoFilter  = false
	Dim filter
	filter		= Request("filter")
	If filter = "ja" then
		autoFilter = true
	End if

	' Wenn Autofilter aktiv d.h. wenn Anwender einen Filter auswählt
	Dim filterText
	Dim spalte
	Dim filterAktiv
	filterText	= ChangeHTMLInUml2(Request("filterText"))
	spalte		= Request("spalte")

	' Farbliche Hervorhebung der aktiven Filter-Spalte
	Dim aktivColor
	aktivColor="#EEEEEE"
	' Abfragen, ob Filter aktiv (Variablen gefüllt) und enstprechende Variable setzen
	if filterText <> "" AND spalte <> "" then
		filterAktiv = true
	end if
	' ------------------------------------- [/AUTOFILTER] ------------------------------------




	' -------------------------------------  [TEXTSUCHE] ------------------------------------
	Dim textSuche
	textSuche = false
	Dim suche
	suche = Request("suche")
	If suche = "ja" then
		textSuche = true
	End if

	' Wenn Textsuche aktiv d.h. wenn Anwender einen Text sucht
	Dim suchText
	Dim suchSpalte
	Dim textSucheAktiv
	suchText	= ChangeHTMLInUml2(Request("suchText2"))
	suchSpalte  = Request("spalte")
	' Farbliche Hervorhebung der aktiven Filter-Spalte
	aktivColor="#EEEEEE"
	' Abfragen, ob Filter aktiv (Variablen gefüllt) und enstprechende Variable setzen
	if suchText <> "" AND spalte <> "" then
		textSucheAktiv = true
	end if
	' ------------------------------------- [/TEXTSUCHE] ------------------------------------



	' -----------------------------------[ANZAHL DATENSÄTZE] --------------------------------
	Dim newMaxRow
	newMaxRow = session("historylength")
	If Request("newMaxRow") <> "" then
		session("historylength") = Request("newMaxRow")
		newMaxRow = Request("newMaxRow")
	End If

	Dim anzahlDatensaetze
	anzahlDatensaetze = newMaxRow

	Dim intMaxRecord
	intMaxRecord=0
	intMaxRecord=Request("intMaxRecord")
	' -----------------------------------[/ANZAHL DATENSÄTZE] --------------------------------


	' -----------------------------------[ MONAT-JAHRES AUSWAHL ] ----------------------------
	Dim monatParam
	Dim jahrParam

	If Request("monat") <> "" then
		monatParam = Request("monat")
	Else
		monatParam = "-"
	End If

	If Request("jahr") <> "" then
		jahrParam = Request("jahr")
	Else
		jahrParam = "-"
	End If
	' -----------------------------------[ /MONAT-JAHRES AUSWAHL ] ----------------------------



	'-------------------------------- [ÜBERGABEPARAMETER AUSWERTEN] --------------------------
	Dim meldeArt2
	Dim meldeKat2
	Dim datum
	datum = ""

	meldeArt2 	= Request("meldeart")
	meldeKat2 	= Request("kategorie")
	datum		= Request("datum")

 	Dim von,bis
 	von = Request("von")
 	bis = Request("bis")

 	Dim sql
 	sql = ChangeHTMLInUml2(Request("sql"))

 	Dim bTopDown
 	bTopDown = Request("bTopDown")

    call setUAdrColumnNameActive(Session("ConnectionString"))

 	'-------------------------------- [/ÜBERGABEPARAMETER AUSWERTEN] --------------------------

%>
<body onload="openFile();">

	<%
	' --------------------------------- [ DATEI ERSTELLEN ] -----------------------------------
	' ST dateiName = replace(date,".","_") &"-"& replace(replace(time,":","_"),".","_") & ".csv"
	' ST dateiName = replace(dateiName,"/","_")
	dateiName = "ActStoring.csv"
	dateiName2 = "temp/"& dateiName ' Dateiname für Öffnen im Browser (Excel)
	' ST tempPfad = Server.MapPath("./")& "\temp"
	tempPfad = "c:\Install\SysteembeheerST"

	' ALte Tempfiles aus Ordner löschen
	deleteOldTempFiles tempPfad

	' Existiert Temp-Ordner? --> Wenn nicht, anlegen!
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (Not fso.FolderExists(tempPfad)) Then
		Set f = fso.CreateFolder(tempPfad)
	End If
	dateiName = tempPfad &"\"& dateiName
	'Response.Write dateiName

	Set b  =  fso.CreateTextFile(dateiName,true)' Datei öffnen
	' --------------------------------- [ /DATEI ERSTELLEN ] -----------------------------------
	
	Dim txtLineString
	'============================================== [ HISTORISCHE ALARME ] =============================================
	if meldeArt2 = "1" then 'Historische Alarme
        aktDatum = datum
		' ------------------------------------- [ CSV-KOPFZEILE ] ---------------------------------
		txtLineString = ben022 &";"& ben023 &";"& ben024 &";"

		' ST If Session("optUseradr") = 1 then
			txtLineString = txtLineString & ben027 & ";"
		' ST End If

		txtLineString = txtLineString & ben025 &";" & ben081 &";"& ben026 &";"


	' ------------------------------------ [ /CSV-KOPFZEILE ] ---------------------------------

	'============================================== [ /HISTORISCHE ALARME ] ==============================================
	else

' <!--============================================== [ AKTIVE/QUITTIERTE ALARME ] ==========================================-->

		txtLineString = ben022 &";"& ben023 &";"& ben024 &";"& ben025 &";" & ben081 & ";"

		' ST If Session("optUseradr") = 1 then
			txtLineString = txtLineString & ben027 & ";"
		' ST End If
	End if

	' Kopfzeile in Datei schreiben
	'Response.Write txtLineString
	b.WriteLine(txtLineString)
	txtLineString=""


'<!--============================================== [ /AKTIVE-QUITTIERTE ALARME ] ==========================================-->

	On Error Resume Next
	Dim currentRecord
	currentRecord = 0

	Set Connection = Server.CreateObject("ADODB.Connection")

	'============================================== [ HISTORISCHE ALARME ] ==============================================
	if meldeArt2 = "1" then 'Historische Alarme

		Dim conString
		Dim dbq

		' Connectionstring zusammenbauen
		
		dbq=Session("Projektpfad")&"\alarm\ProtBACnet\"&aktDatum&".mdb"
		conString= "Driver={Microsoft Access Driver (*.mdb)}; Dbq="&dbq&"; Uid=; Pwd=;"
		
		' Verbindung herstellen
		Connection.Open conString
		'Response.Write sql
		if Err.Number = 0 then 'Wenn Datenbankverbindung ohne Fehler
			On Error Resume Next
			Dim bNoRecords
			Set rs2 = CreateObject ("ADODB.Recordset")
			rs2.CursorType = adOpenKeyset
			rs2.LockType = adLockOptimistic

			rs2.Open sql, Connection
			If Not rs2.EOF and Not rs2.BOF Then
				avarRecords = rs2.GetRows()
			Else
				bNoRecords = True
			End If
			'rs2.MoveFirst
			'rs2.MoveLast
			rs2.Close
			Dim nPrevPaging,nLastPaging
			Set rs2 = Nothing
			If Not bNoRecords Then
				nRecords = UBound(avarRecords,2)
				'Response.Write "RECORD "&nRecords
				If Not bTopDown Then
					nPrevPaging = avarRecords(0,0)
					nLastPaging = avarRecords(0,nRecords)
				Else
					nPrevPaging = avarRecords(0,nRecords)
					nLastPaging = avarRecords(0,0)
				End If
			End If

			Dim recordAnzCount
			recordAnzCount=0
			If Not bTopDown Then
				For intRecord = 0 To nRecords

					exportRecord avarRecords, intRecord
					currentRecord=currentRecord+1
					'Response.Flush
				Next
			Else
				For intRecord = nRecords To 0 Step -1
					exportRecord avarRecords, intRecord
					currentRecord=currentRecord+1
					'Response.Flush
				Next
			End If
			' ----------------------------------- [/NAVIGATION] --------------------------------------------
		End If
		'--------------------------------- [TAGESANSICHT] -------------------------------------
		'b.close
	'============================================== [ /HISTORISCHE ALARME ] ==============================================


	'========================================= [ AKTUELLE / QUITTIERTE ALARME ] ==============================================
	else 'Alle anderen Alarme (aktive/quittierte)
		'--------------------------------- [CONNECTION AUFBAUEN] -------------------------------------
		Connection.Open Session("ConnectionString")
		SQLStmt = "SELECT Meldeart, "& userAdr_SpaltenName &", UserAdr, Lieg, DT_Stamp_high, PlantID, Alarmtext, " & _
            " Path, Ust, Reg, quittime, quit_user, prio, LogTime, Sondertextdatei, dv_path, Protokoll, " & _
            " PhyAdresse, AckedTrans, BACnetStatus, quitted  FROM Sondertexte" & _
            " WHERE Meldeart=" &CInt(meldeKat2)& ""
        if Not filterAktiv then 'Wenn Filter nicht aktiv
            If textSucheAktiv then
                If UCase(spalte) = "PRIO" then
                    If meldeArt2 = "255" then 'gesperrte Meldungen
                        SQLStmt = SQLStmt & " AND "&spalte&" = "&suchText&" AND prio BETWEEN 250 AND " & _
                            CInt(meldeArt2) &" AND trim(alarmtext)<>'_' ORDER BY LogTime DESC"
                    Else
                        SQLStmt = SQLStmt & " AND "&spalte&" = "&suchText&" AND quitted="& CInt(meldeArt2) & _
                            " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND trim(alarmtext)<>'_' ORDER BY DT_Stamp_high DESC"
                        'Response.Write "SQLStmt: " & SQLStmt
                    End If
                Else
                    If meldeArt2 = "255" then 'gesperrte Meldungen
                        SQLStmt = SQLStmt & " AND "&spalte&" like '%"&suchText&"%' AND prio BETWEEN 250 AND " & _
                            CInt(meldeArt2) &" AND trim(alarmtext)<>'_' ORDER BY LogTime DESC"
                    Else ' All active alarms
                        If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                            SQLStmt = SQLStmt & " AND "&spalte&" like '%" & suchText & "%' AND high_low=-1 AND prio " & _
                                "BETWEEN 1 AND 250 AND trim(alarmtext)<>'_' ORDER BY DT_Stamp_high DESC"
                        Else
                            SQLStmt = SQLStmt & " AND "&spalte&" like '%"&suchText&"%' AND quitted="& CInt(meldeArt2) & _
                                " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND trim(alarmtext)<>'_' ORDER BY DT_Stamp_high DESC"
                        End If
                    End If
                End If
            Else
                If meldeArt2 = "255" then 'gesperrte Meldungen
                    SQLStmt = SQLStmt & " AND prio BETWEEN 250 AND "& CInt(meldeArt2) & _
                        " AND trim(alarmtext)<>'_' ORDER BY LogTime DESC"
                Else
                    If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                        SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND trim(alarmtext)<>'_' " & _
                            "ORDER BY DT_Stamp_high DESC"
                    Else
                        SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND prio BETWEEN 1 " & _
                            "AND 250 AND trim(alarmtext)<>'_' ORDER BY DT_Stamp_high DESC"
                    End If
                End If
            End If
        else ' Wenn Filter aktiv
            If meldeArt2 = "255" then 'gesperrte Meldungen
                SQLStmt = SQLStmt & " AND "&spalte&"='"&filterText&"' AND prio BETWEEN 250 AND "& CInt(meldeArt2) & _
                    " AND trim(alarmtext)<>'_' ORDER BY LogTime DESC"
            Else
                If UCase(spalte) = "DT_STAMP_HIGH" then 'Bei Datumsfilter, Datum und SQL-Statement anpassen.
                    Dim datum4
                    datum4 = CDate(filterText)
                    datum4 = month(datum4)&"/"&day(datum4)&"/"&year(datum4)&" "& FormatDateTime(datum4,3)
                    If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                        SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250" & _
                            " AND trim(alarmtext)<>'_' AND " & spalte & "=#" & datum4 & _
                            "# ORDER BY DT_Stamp_high DESC"
                    Else
                        SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND " & _
                            "prio BETWEEN 1 AND 250 AND trim(alarmtext)<>'_' AND " & spalte & "=#" & datum4 & _
                            "# ORDER BY DT_Stamp_high DESC"
                    End If
                Else
                    If UCase(spalte) = "PRIO" then ' If selected autofilter column is "prio"
                        If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                            SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND" & _
                            " trim(alarmtext)<>'_' AND " & spalte & "=" & filterText & _
                                " ORDER BY DT_Stamp_high DESC"
                        Else
                            SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND " & _
                                "prio BETWEEN 1 AND 250 AND trim(alarmtext)<>'_' AND " & spalte & "=" & filterText & _
                                " ORDER BY DT_Stamp_high DESC"
                        End If
                    Else
                        If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                            SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND" & _
                            " trim(alarmtext)<>'_' AND " & spalte & "='" & filterText & _
                                "' ORDER BY DT_Stamp_high DESC"
                        Else
                            SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND " & _
                                "prio BETWEEN 1 AND 250 AND trim(alarmtext)<>'_' AND " & spalte & "='" & filterText & _
                                "' ORDER BY DT_Stamp_high DESC"
                        End If
                    End If
                End If
            End If
        end if

		'--------------------------------- [/CONNECTION AUFBAUEN] -------------------------------------


		'-------------------------------------- [Write file]-------------------------------------------
		%> <!-- ST -->
		<p> { <br/> <!-- ST -->
		<% ' ST
		Set RS = Connection.Execute(SQLStmt)

		Do While CheckRS(RS)

			' Ist die Liegenschaft des Alarms in dem Liegenschaftsstring des User vorhanden...?
			' ST if instr(Session("LiegenschaftSet"),Replace(CheckRequest(RS, "Lieg")," ","") & ";") > 0 OR instr("WEBVISION",CheckRequest(RS, "Lieg")) > 0 then

				' Abfangen ob Alarm gesperrt dann Logdatei anzeigen
				If meldeArt2 = "255" then
					txtLineString= CheckRequest(RS, "LogTime") &";"
				Else
					txtLineString= CheckRequest(RS, "DT_Stamp_high") &";"
				End If

				txtLineString = txtLineString & trim(replace(CheckRequest(RS, "Alarmtext"),"_"," ")) &";"
				txtLineString = txtLineString & trim(replace(CheckRequest(RS, "Lieg"),"_"," ")) &";"
				txtLineString = txtLineString & trim(replace(CheckRequest(RS, "PlantID"),"_"," ")) &";"
                txtLineString = txtLineString & trim(CheckRequest(RS, "prio") & " (" & arrPriority(CInt(RS("prio").value)) & ")") & ";"

				' ST If Session("optUseradr") = 1 then
					txtLineString = txtLineString & trim(replace(CheckRequest(RS, "UserAdr"),"_"," ")) &";"
				' ST End If
				b.WriteLine(txtLineString)
				%> <!-- ST -->
				<%=txtLineString%><br/> <!-- ST -->
				<% ' ST
				txtLineString=""
				currentRecord=currentRecord+1

			' ST end if
			RS.MoveNext
		Loop
		%> <!-- ST -->
		{ 
		<% ' ST
		'-------------------------------------------- [/Write file]-------------------------------------------


	End if
    ' Cleaning-up operations
	b.close
    Connection.Close
    Set Connection = nothing
    Set RS  = nothing
    Set rs2 = nothing
    Set fso = nothing
    Set b   = nothing
    Set f   = nothing	
			
	%>



</body>
<script language="JavaScript" type="text/javascript">
    // ST Link to file for open download window
    function openFile() {
         window.location.href = "<%=dateiName%>";
    }

    // ST After the download operation beginns to start, go back to the latest view,
    // ST or, in case of actual events, close the window.
    // ST function goBack() {
    // ST    <% If meldeArt2 = "1" then 'Historische Alarme %>
    // ST    window.setTimeout("history.go(-1);", 1000);
    // ST    <% Else %>
    // ST    window.setTimeout("window.close()", 1000);
    // ST   <% End If %>
    // ST }
</script>
</html>