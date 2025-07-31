<%@ LANGUAGE="VBscript" %>
<!--#include virtual="/webvisionnt/zsystem/includes/doctype.asp"-->

<%
' ###################################################################################################
' Sprachsteuerungs - Session holen und dementsprechende Sprachdatei
' einlesen und ausführen (Konstanten registrieren/deklarieren...).
On Error Resume Next
Dim lang
lang = Session("Language")
if lang = "" then
    lang = "DE"
end if

Const HIST_PATH = "\alarm\ProtBACnet"
Dim historyPath : historyPath = Session("Projektpfad") & HIST_PATH

Dim uStyle
If Session("userStyle")<>"" then ustyle=trim(Session("userStyle")) else ustyle="s3" End If

'--Zwangsabmeldecounter zurücksetzen durch aktion---
Session("inactivetime") = 0
'---------------------------------------------------

' ###################################################################################################
%>

<!--#include virtual="/webvisionnt/zsystem/lang/Alarm/texteAlarm.asp"-->
<!--#include file="stringFunktionen.asp"-->
<!--#include file="ADOVBS.inc"-->
<!--#include file="IASUtil.asp"-->
<!--#include file="WEB.inc"-->
<!--#include file="WEB2.inc"-->
<!--#include file="priorityDescription_include.asp"-->
<html>
<head>
<script language="Javascript" type="text/javascript">
    var quittA  = new Image();
    quittA.src  = "../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_A.gif";
    var quittF  = new Image();
    quittF.src  = "../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_F.gif";
    var quittT  = new Image();
    quittT.src = "../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_T.gif";
</script>

<link rel="Stylesheet" type="text/css" href="../styles/<%=uStyle%>/schriften.css" />
<title></title>
</head>
<%
    'On Error Resume Next
    Response.Buffer = true

    call setUAdrColumnNameActive(Session("ConnectionString"))

    ' ------------------------------------- [AUTOFILTER] -------------------------------------
    ' Boolean um festzustellen ob Seite mit Autofilter geladen wird
    Dim autoFilter : autoFilter = false
    Dim filter : filter = Request.QueryString("filter")
    If filter = "ja" then
        autoFilter = true
    End if

    ' Wenn Autofilter aktiv d.h. wenn Anwender einen Filter auswählt
    Dim filterText : filterText = ChangehtmlInUml2(Request.QueryString("filterText"))
    Dim spalte : spalte = Request.QueryString("spalte")
    Dim filterAktiv
    
    ' Farbliche Hervorhebung der aktiven Filter-Spalte
    Dim aktivColor : aktivColor = "bgRowColor2"
    ' Abfragen, ob Filter aktiv (Variablen gefüllt) und enstprechende Variable setzen
    if filterText <> "" AND spalte <> "" then
        filterAktiv = true
    end if
    ' ------------------------------------- [/AUTOFILTER] ------------------------------------



    ' -------------------------------------  [TEXTSUCHE] ------------------------------------
    Dim textSuche : textSuche = false
    Dim suche : suche = Request.QueryString("suche")
    If suche = "ja" then
        textSuche = true
    End if

    ' Wenn Textsuche aktiv d.h. wenn Anwender einen Text sucht
    Dim suchText : suchText = ChangehtmlInUml2(Request.QueryString("suchText2"))
    Dim suchSpalte : suchSpalte  = Request.QueryString("spalte")
    Dim textSucheAktiv
    
    ' Abfragen, ob Filter aktiv (Variablen gefüllt) und enstprechende Variable setzen
    if suchText <> "" AND spalte <> "" then
        textSucheAktiv = true
    end if
    ' ------------------------------------- [/TEXTSUCHE] ------------------------------------



    ' -----------------------------------[ANZAHL DATENSÄTZE] --------------------------------
    ' Maximale Anzahl anzuzeigender Datensätze auf einer html-Seite
    ' Voreinstellung i.d.R. 300 Rows
    Dim newMaxRow : newMaxRow = Session("historylength")
    
    ' Wenn Seite mit einer neuen Anzahl von Datensätzen aufgerufen wird...
    If Request.QueryString("newMaxRow") <> "" then
        ' ...diese in der Sessionvariablen speichern...
        session("historylength") = Request.QueryString("newMaxRow")
        ' ...und lokal im script speichern.
        newMaxRow = Request.QueryString("newMaxRow")
    End If
    Dim anzahlDatensaetze : anzahlDatensaetze = newMaxRow
    ' -----------------------------------[/ANZAHL DATENSÄTZE] --------------------------------


    ' -----------------------------------[ MONAT-JAHRES AUSWAHL ] ----------------------------
    Dim monatParam
    Dim jahrParam
    If Request.QueryString("monat") <> "" then
        monatParam = Request.QueryString("monat")
    Else
        monatParam = "-"
    End If

    If Request.QueryString("jahr") <> "" then
        jahrParam = Request.QueryString("jahr")
    Else
        jahrParam = "-"
    End If
    ' -----------------------------------[ /MONAT-JAHRES AUSWAHL ] ----------------------------



    '-------------------------------- [ÜBERGABEPARAMETER AUSWERTEN] --------------------------
    Dim sortOrder : sortOrder = Request.QueryString("sort")
    Dim meldeArt2 : meldeArt2 = Request.QueryString("meldeart")
    Dim meldeKat2 : meldeKat2 = Request.QueryString("kategorie")
    Dim datum : datum = Request.QueryString("datum")

    ' Session für aktFrame füllen (Keine Aktualisierung bei Ansicht der Historie)
    Session("MeldeKategorie") = meldeArt2

    Dim ueberschrift, uKategorie
    Dim tableCellClassName
    Dim tablHeaderClassName
    ' Get the event description from the event category
     Select Case meldeKat2
        case 0
            uKategorie= ben002
        case 1
            uKategorie= ben002
        case 2
            uKategorie= ben003
        case 3
            uKategorie= ben004
    End Select
    'uKategorie = getEventText(meldeKat2)

    ' Tabellen - Überschrift formatieren
    if meldeArt2 = "0"  then 'Aktive Alarme
        ueberschrift = ben009 & "&nbsp;&nbsp;" & uKategorie
        If uKategorie = ben002 then
            tableCellClassName = "actAlarmTableCell" ' Bei Alarmeldung Schriftfarbe auf Rot setzen...
            tablHeaderClassName = "actAlarmTableHeader"
        else
            tableCellClassName = "inActAlarmTableCell" ' .... ansonsten Schwarz
            tablHeaderClassName = "inActAlarmTableHeader"
        End If
    elseif meldeArt2 = "-255" then ' Aktive UND Quitierte Alarme
        ueberschrift = ben009 & "&nbsp;+&nbsp;" & ben009 & "&nbsp;" & ben007 & "&nbsp;&nbsp;" & uKategorie
        tableCellClassName = "actAlarmTableCell" ' Bei Alarmeldung Schriftfarbe auf Rot setzen...
        tablHeaderClassName = "actAlarmTableHeader"
    elseif meldeArt2 = "-1" then ' Quittierte Alarme
        ueberschrift = ben009 & "&nbsp;&nbsp;" & ben007 & "&nbsp;&nbsp;" & uKategorie
        tableCellClassName = "inActAlarmTableCell"
        tablHeaderClassName = "inActAlarmTableHeader"
    elseif meldeArt2 = "1" then 'Historische Alarme
        ueberschrift = ben010 & "&nbsp;&nbsp;" & uKategorie
        tableCellClassName = "inActAlarmTableCell"
        tablHeaderClassName = "inActAlarmTableHeader"
    elseif meldeArt2 = "255" then 'Gesperrte Alarme
        ueberschrift = ben012 & "&nbsp;&nbsp;" & uKategorie
        tableCellClassName = "inActAlarmTableCell"
        tablHeaderClassName = "inActAlarmTableHeader"
    End if

    ' ST Dim qUserRight, qUserSession
    ' ST qUserSession = CStr(Session("optQUser"))
    ' ST If qUserSession <> "" AND qUserSession = "1" then
        qUserRight = true
    ' ST Else
    ' ST     qUserRight = false
    ' ST End If

    ' Tabellenspalten anpassen wenn Userdaresse hinzugefügt oder entfernt wird
    Dim colspan
    if Session("optUseradr") = 1 then
        If qUserRight AND meldeArt2 = "-1" then
            colspan = "10"
        Else
            colspan = "9"
        End If
    else
        If qUserRight AND meldeArt2 = "-1" then
            colspan = "9"
        Else
            colspan = "8"
        End If
    end if
    'If meldeArt2 <> "1" then
        colspan = colspan +1
    'End If
    If meldeArt2 = "1" then
        colspan = colspan +1
    End If
    If meldeArt2 = "-255" then
        colspan = colspan +1
    End If
    '-------------------------------- [/ÜBERGABEPARAMETER AUSWERTEN] --------------------------

%>

<body onload="setHistory('<%=meldeArt2%>','<%=meldeKat2%>'); RefreshLogo();" style="padding:10px;">
<%
' Hier noch weiter überlegen/spielen wie die Ermittlung der Wochen eines Monats funktionieren könnte.
' hilfreiche Quelle hierbei: www.asphelper.de
'Dim kwDatum
'kwDatum = CDate("01.03.2011")
'Response.Write("KW:" & KW(kwDatum))
'Response.Write("Weekday 01.03.2011: " & Weekday(kwDatum, 1))
'Response.Write("DateAdd 01.03.2011: " & DateAdd("ww",1, kwDatum))
%>
<table style="width:100%" class="TableOne" cellspacing="1">
    <%
If textSuche then 'Bei Textsuche Form-Tag einfügen mit Abschickung zur textSuche.asp
    %>
    <form method="post" action="textSuche.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&monat=<%=monatParam%>&jahr=<%=jahrParam%>">
    <%
End if
%>
<tr class="bgRowColor2">
    <% If meldeArt2 = "1" then 'Historische Alarme' %>
        <% If Session("optUseradr") = 1 then %>
            <% If qUserRight AND meldeArt2 = "-1" then %>
                <td colspan="<%=colspan%>" class="<%=tablHeaderClassName%>">
            <% Else %>
                <td colspan="<%=colspan%>" class="<%=tablHeaderClassName%>">
            <% End If %>
        <% else %>
            <% If qUserRight AND meldeArt2 = "-1" then %>
                <td colspan="<%=colspan%>" class="<%=tablHeaderClassName%>">
            <% Else %>
                <td colspan="<%=colspan%>" class="<%=tablHeaderClassName%>">
            <% End If %>
        <% End If %>

    <% Else %>
        <td colspan="<%=colspan%>" class="<%=tablHeaderClassName%>">
    <% End If %>
        <%=ueberschrift%>
    </td>
</tr>
<tr class="bgRowColor1">
    <% If meldeArt2 = "1" then 'Historische Alarme' %>
        <% If Session("optUseradr") = 1 then %>
            <% If qUserRight AND meldeArt2 = "-1" then %>
                <td colspan="<%=colspan%>">
            <% Else %>
                <td colspan="<%=colspan%>">
            <% End If %>
        <% else %>
            <% If qUserRight AND meldeArt2 = "-1" then %>
                <td colspan="<%=colspan%>">
            <% Else %>
                <td colspan="<%=colspan%>">
            <% End If %>
        <% End If %>
    <% Else %>
        <td colspan="<%=colspan%>">
    <% End If %>
    <% ' -------------------------------------------[ HISTORISCHE OPTIONEN ] --------------------------------
    If meldeArt2 = "1" then 'Historische Alarme
        ' Aktuelles Datum splitten
        Dim datumArr, aktDatum
        datumArr = Split(nowDeutsch,".")
        if datum = "" then  ' Wenn kein Datum bei Aufruf der Seite übergeben wurde...
            datumArr = Split(nowDeutsch,".") '... aktuelles Datum verwenden,...
            aktDatum = datumArr(2)&""&datumArr(1)
        else ' ...bei rekursiven Aufruf durch Select-Box...
            aktDatum = datum '...übergebenes Datum verwenden.
        End if

        ' 1. Ordner der Historischen Alarme durchsuchen, mdb´s auslesen und Tabellennamen in Array speichern
        call getHistoryDBs(historyPath)
       
    %>
        <input type="button" value="<%=ben015%>" onclick="loadNewMaxRows();" />&nbsp;&nbsp;
        <input style="width:40px;" name="historylength" id="historylength" type="text" value="<%=session("historylength")%>" /> 
        <%=ben016%>

    <% ' -------------------------------------------[ /HISTORISCHE OPTIONEN ] --------------------------------
    End If%>

    <%
    If meldeArt2 <> "1" then
        '------------------------------------------ [TEXTSUCHE AKTIVE / QUITTIERTE ALARME] ------------------------------------------
        If textSuche then 'Wenn der Benutzer den "Textsuche" - Button gedrückt hat
            Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"& ben020 & ":&nbsp;&nbsp;<input type=""text"" name=""textSuche"" value=""" & suchText & """><input type=""Image"" src=""../styles/"& uStyle &"/bin/alarme/suchPfeil.gif"" style=""cursor:hand;"">"
        End if
        '------------------------------------------ [/TEXTSUCHE AKTIVE / QUITTIERTE ALARME] ------------------------------------------
    End If
    %>
    </td>
</tr>
    <%
    '============================================== [ HISTORISCHE ALARME ] ==============================================
    if meldeArt2 = "1" then 'Historische Alarme

        '----------------------------------- [SELECTBOX FÜLLEN (MONATE)] -------------------------------

        ' 2. Array durchgehen und Jahreszahlen in Selectbox anzeigen
        %>
        <tr class="bgRowColor1" >
            <% If Session("optUseradr") = 1 then %>
                <% If qUserRight AND meldeArt2 = "-1" then %>
                    <td colspan="12" style="vertical-align:bottom;" valign="bottom">
                <% Else %>
                    <td colspan="11" style="vertical-align:bottom;" valign="bottom">
                <% End If %>
            <% else %>
                <% If qUserRight AND meldeArt2 = "-1" then %>
                    <td colspan="11" style="vertical-align:bottom;" valign="bottom">
                <% Else %>
                    <td colspan="10" style="vertical-align:bottom;" valign="bottom">
                <% End If %>
            <% End If %>
            <%=ben018%>:&nbsp;&nbsp;<select name="zeit" size="1" onchange="changeDay(this.value)" style="width:150px;">
            <option value="">--<%=ben019%>--</option>
            <%
            Dim selectedD
            selectedD =""
            Dim zeitString
            Dim aktivMonat, monat, jahr
            Dim i
            If isArrayFilled(dbTimeArray) then
                For i = 0 to UBound(dbTimeArray)
                    'Response.Write dbTimeArray(i)
                    ' Monat und Jahr ausfiltern
                    monat = Mid(dbTimeArray(i),5,2)
                    jahr = left(dbTimeArray(i),4)
                    zeitString = jahr&"-"&monat&" / "&MName(monat,ben032,ben033,ben034,ben035,ben036,ben037,ben038,ben039,ben040,ben041,ben042,ben043)
                    ' Eintrag vorselektieren
                    'Response.Write "DATUM: " & aktDatum
                    if aktDatum = dbTimeArray(i) then
                        selectedD = "selected"
                        aktivMonat = Mid(dbTimeArray(i),5,2)
                    End if
                    Response.Write "<option value="""&dbTimeArray(i)&""" "&selectedD&">"&zeitString&"</option>"&vbCrLf
                    selectedD=""
                Next
            End If
            Response.Write "</select>"&vbCrLf
            if filterAktiv then
                Response.Write "<span style=""color:#00FF7F; text-align:right; font-weight:bold;"">&nbsp;"& ben053 &"</span>"
            End if

            '-------------------------------------------------------------------------
            ' Spaltennamen für Useradresse
            userAdr_SpaltenNameH = "C_Adr"
            '-------------------------------------------------------------------------

            '------------------------------------------ [TEXTSUCHE] ------------------------------------------
            If textSuche then
                Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"& ben020 &":&nbsp;&nbsp;<input type=""text"" name=""textSuche"" value="""&suchText&"""><input type=""Image"" src=""../styles/"& uStyle &"/bin/alarme/suchPfeil.gif"" style=""cursor:hand;"">"
            End if
            '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
            %>
            <img src="../styles/<%=uStyle%>/bin/alarme/prev.gif" alt="<%=ben054%>" border="0" style="vertical-align:bottom; cursor:hand; position:absolute; right:35px;" onclick="goWay('Prev')">
            <img src="../styles/<%=uStyle%>/bin/alarme/next.gif" alt="<%=ben055%>" border="0"  style="vertical-align:bottom; cursor:hand; position:absolute; right:10px;" onclick="goWay('Next')">
            <%

            '----------------------------------- [/SELECTBOX FÜLLEN (MONATE)] -------------------------------
            %>
            </td>
        </tr>

        <tr class="bgRowColor2">
            <%
            Dim tValue(6)
            tValue(0) = "TStamp"
            tValue(1) = "Messagetext"
            tValue(2) = "Liegenschaft"
            tValue(3) = userAdr_SpaltenNameH
            tValue(4) = "Plant"
            tValue(5) = "Ereignistext"
            tValue(6) = "Priority"

            Dim checked
            checked=""
            if spalte = tValue(0) then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
            %>
            <td><img src="../styles/<%=uStyle%>/bin/alarme/quittA.gif" alt="<%=ben151 %>" title="<%=ben151 %>" /></td>
            <td class="<%=aktivColor%>" align="center" nowrap = "nowrap">
                <%
                    '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                    If textSuche then
                        %>
                        <input type="radio" value="<%=tValue(0)%>" name="OptionBox" <%=checked%> />
                        <%
                    End if
                    '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
                %>
                <b><%=ben022%></b>
                <!--
                <a href="<% = getSortURLAdress(tValue(0) & " ASC")%>">
                    <img border="0" src="../styles/<%=uStyle%>/bin/up.gif" alt="<%=ben171%>" title="<%=ben171%>" />
                </a>
                <a href="<% = getSortURLAdress(tValue(0) & " DESC")%>">
                    <img border="0" src="../styles/<%=uStyle%>/bin/down.gif" alt="<%=ben172%>" title="<%=ben172%>" />
                </a>-->
            </td>
            <%
            aktivColor="bgRowColor2"
            checked=""
            if textSucheAktiv = false AND textSuche = true then
                checked="checked"
            End if
            if spalte = tValue(1) then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
            %>
            <td class="<%=aktivColor%>" align="center">
                <%
                    '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                    If textSuche then
                    %>
                    <input type="radio" value="<%=tValue(1)%>" name="OptionBox" <%=checked%> />
                    <%
                    End if
                    '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
                %>
                <b><%=ben023%></b>
            </td>
            <%
            aktivColor="bgRowColor2"
            checked=""
            if spalte = tValue(2) then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
            %>

            <td class="<%=aktivColor%>" align="center">
                <%
                    '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                    If textSuche then
                    %>
                    <input type="radio" value="<%=tValue(2)%>" name="OptionBox" <%=checked%> />
                    <%
                    '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
                    End if
                %>
                <b><%=ben024%></b>
            </td>


            <%
            If Session("optUseradr") = 1 then
                aktivColor="bgRowColor2"
                checked=""
                if spalte = tValue(3) then
                    aktivColor="bgActiveColor"
                    checked="checked"
                End if
                %>
                <td class="<%=aktivColor%>" align="center">
                    <%
                        '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                        If textSuche then
                        %>
                        <input type="radio" value="<%=tValue(3)%>" name="OptionBox" <%=checked%> />
                        <%
                        End if
                        '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
                    %>
                    <b><%=ben027%></b>
                </td>
            <%
            End If

            aktivColor="bgRowColor2"
            checked=""
            if spalte = tValue(4) then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
            %>
            <td class="<%=aktivColor%>" align="center">
                <%
                    '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                    If textSuche then
                    %>
                    <input type="radio" value="<%=tValue(4)%>" name="OptionBox" <%=checked%> />
                    <%
                    End if
                    '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
                %>
                <b><%=ben025%></b>
            </td>
            <%
            aktivColor="bgRowColor2"
            checked=""
            if spalte = tValue(6) then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
            %>
            <td class="<%=aktivColor%>" align="center">
                <%
                    '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                    If textSuche then
                    %>
                    <input type="radio" value="<%=tValue(6)%>" name="OptionBox" <%=checked%> />
                    <%
                    End if
                    '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
                %>
                <b><%=ben081%></b>
            </td>

            <%
            aktivColor="bgRowColor2"
            checked=""
            if spalte = tValue(5) then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
            %>
            <td class="<%=aktivColor%>" align="center" nowrap>
                <%
                    '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                    If textSuche then
                    %>
                    <input type="radio" value="<%=tValue(5)%>" name="OptionBox" <%=checked%> />
                    <%
                    End if
                    '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
                %>
                <b><%=ben026%></b>
            </td>
            <td colspan="3">
            </td>
            <%
                aktivColor="bgRowColor2"
                checked=""
            %>
    <%
        '============================================== [ /HISTORISCHE ALARME ] ==============================================
        else ' aktive Alarme
    %>
    <tr class="bgRowColor1">
        <%
            ' Bei Textsuche Such-Spalte grün markieren und passende Option-Box auswählen
            aktivColor="bgRowColor3"
            If meldeArt2 = "255" then
                if spalte = "LogTime" then
                    aktivColor="bgActiveColor"
                    checked="checked"
                End if
            Else
                If spalte = "date_time_stamp_high" then
                    aktivColor="bgActiveColor"
                    checked="checked"
                End if
            End If
        %>
        <td></td>
        <td align="center" class="<%=aktivColor%>" nowrap = "nowrap">
        <%
            '------------------------------------------ [TEXTSUCHE] ------------------------------------------
            Dim sortOrderColumnName : sortOrderColumnName = "DT_Stamp_high"
            If textSuche then
                If meldeArt2 = "255" then
                    sortOrderColumnName = "LogTime"
                %>
                    <input type="radio" value="LogTime" name="OptionBox" <%=checked%> />
                <%
                Else
                    sortOrderColumnName = "DT_Stamp_high"
                %>
                    <input type="radio" value="DT_Stamp_high" name="OptionBox" <%=checked%> />
                <%
                End If
            End if
            '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
        %>
            <b><%=ben022%></b>
            <a href="<% = getSortURLAdress(sortOrderColumnName & " ASC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/up.gif" alt="<%=ben171%>" title="<%=ben171%>" />
            </a>
            <a href="<% = getSortURLAdress(sortOrderColumnName & " DESC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/down.gif" alt="<%=ben172%>" title="<%=ben172%>" />
            </a>
        </td>
        <%
            aktivColor="bgRowColor3"
            checked=""
            ' Optionbox auf Alarmtext vorselektieren
            if textSucheAktiv = false AND textSuche = true then
                checked="checked"
            End if

            if spalte = "Alarmtext" then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
        %>
        <td align="center" class="<%=aktivColor%>">
        <%
            '------------------------------------------ [TEXTSUCHE] ------------------------------------------
            If textSuche then
            %>
            <input type="radio" value="Alarmtext" name="OptionBox" <%=checked%> />
            <%
            End if
            '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
        %>
            <b><%=ben023%></b>
            <a href="<% = getSortURLAdress("Alarmtext ASC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/up.gif" alt="<%=ben171%>" title="<%=ben171%>" />
            </a>
            <a href="<% = getSortURLAdress("Alarmtext DESC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/down.gif" alt="<%=ben172%>" title="<%=ben172%>" />
            </a>
        </td>
        <%
            aktivColor="bgRowColor3"
            checked=""
            if spalte = "Lieg" then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
        %>
        <td align="center" class="<%=aktivColor%>">
        <%
            '------------------------------------------ [TEXTSUCHE] ------------------------------------------
            If textSuche then
            %>
            <input type="radio" value="Lieg" name="OptionBox" <%=checked%> />
            <%
            End if
            '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
        %>
            <b><%=ben024%></b>
            <a href="<% = getSortURLAdress("Lieg ASC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/up.gif" alt="<%=ben171%>" title="<%=ben171%>" />
            </a>
            <a href="<% = getSortURLAdress("Lieg DESC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/down.gif" alt="<%=ben172%>" title="<%=ben172%>" />
            </a>
        </td>
        <%
            aktivColor="bgRowColor3"
            checked=""
            if spalte = "PlantId" then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
        %>
        <td align="center" class="<%=aktivColor%>">
        <%
            '------------------------------------------ [TEXTSUCHE] ------------------------------------------
            If textSuche then
            %>
            <input type="radio" value="PlantId" name="OptionBox" <%=checked%> />
            <%
            End if
            '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
        %>
            <b><%=ben025%></b>
            <a href="<% = getSortURLAdress("PlantId ASC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/up.gif" alt="<%=ben171%>" title="<%=ben171%>" />
            </a>
            <a href="<% = getSortURLAdress("PlantId DESC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/down.gif" alt="<%=ben172%>" title="<%=ben172%>" />
            </a>
        </td>
        <%
            aktivColor="bgRowColor3"
            checked=""
            if spalte = "prio" then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
        %>
        <td align="center" class="<%=aktivColor%>">
        <%
            '------------------------------------------ [TEXTSUCHE] ------------------------------------------
            If textSuche then
            %>
            <input type="radio" value="prio" name="OptionBox" <%=checked%> />
            <%
            End if
            '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
        %>
            <b><%=ben081%></b>
            <a href="<% = getSortURLAdress("prio ASC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/up.gif" alt="<%=ben171%>" title="<%=ben171%>" />
            </a>
            <a href="<% = getSortURLAdress("prio DESC")%>">
                <img border="0" src="../styles/<%=uStyle%>/bin/down.gif" alt="<%=ben172%>" title="<%=ben172%>" />
            </a>
        </td>
        <%
            aktivColor="bgRowColor3"
            checked=""
            if spalte = userAdr_SpaltenName then
                aktivColor="bgActiveColor"
                checked="checked"
            End if
        %>
        <% If Session("optUseradr") = 1 then %>
            <td align="center" class="<%=aktivColor%>">
            <%
                '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                If textSuche then
                %>
                <input type="radio" value="<%=userAdr_SpaltenName%>" name="OptionBox" <%=checked%> />
                <%
                End if
                '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
            %>
                <b><%=ben027%></b>
                <a href="<% = getSortURLAdress(userAdr_SpaltenName & " ASC")%>">
                    <img border="0" src="../styles/<%=uStyle%>/bin/up.gif" alt="<%=ben171%>" title="<%=ben171%>" />
                </a>
                <a href="<% = getSortURLAdress(userAdr_SpaltenName & " DESC")%>">
                    <img border="0" src="../styles/<%=uStyle%>/bin/down.gif" alt="<%=ben172%>" title="<%=ben172%>" />
                </a>
            </td>
        <% End If %>


        <%
        aktivColor="bgRowColor3"
        If qUserRight AND meldeArt2 = "-1" then %>
            <td align="center" class="<%=aktivColor%>">
            <%
                '------------------------------------------ [TEXTSUCHE] ------------------------------------------
                If textSuche then
                %>
                <input type="radio" value="<%=userAdr_SpaltenName%>" name="OptionBox" <%=checked%> />
                <%
                End if
                '------------------------------------------ [/TEXTSUCHE] ------------------------------------------
            %>
                <b><%=ben155%></b>
            </td>
        <% End If %>
        <% If meldeArt2 = "-255" then 'Aktive und quittierte Meldungen %>
            <td align="center" style="font-weight:bold;" class="<%=aktivColor%>"><%=ben180 %></td>
            <td align="center" style="font-weight:bold;" colspan="3" class="<%=aktivColor%>">
                <!-- <%=ben183%> -->&nbsp;
            </td>
        <% Elseif meldeArt2 = "-1" then ' Only acknowledged events %>
            <td align="center" style="font-weight:bold;" colspan="3" class="<%=aktivColor%>">
                <!-- <%=ben183%> -->&nbsp;
            </td>
        <% Else %>
            <td align="center" style="font-weight:bold;" class="<%=aktivColor%>"><%=ben180 %></td>
            <td align="center" style="font-weight:bold;" colspan="2" class="<%=aktivColor%>">
                <!-- <%=ben183%> -->&nbsp;
            </td>
        <% End if %>
        <%
            ' Aktiv-Farbe und Option-Box zurücksetzen
            aktivColor="bgRowColor3"
            checked=""
        %>

<% End if %>

</tr>
<% 
If textSuche then '
%>
    </form>
<% End If %>

<%
    On Error Resume Next
    Err.Clear
    Dim currentRecord
    currentRecord = 0
    Dim intMaxRecord
    intMaxRecord=0
    Dim Connection
    Set Connection = Server.CreateObject("ADODB.Connection")

    '============================================== [ HISTORISCHE ALARME ] ==============================================
    if meldeArt2 = "1" then 'Historische Alarme
        %>
            <!-- Formular zum Druck aller aufgelisteten Alarme (Parameterübergabe nur per "Post" möglich, da "Get" zu wenig Zeichen zulässt (Url-Parameter beschränkt auf 2000 Zeichen)) -->
            <form name="Druckansicht" method="post" action="printAlarme.asp" target="_blank">
                <input type="hidden"  name="sql"        value="<%=SQLStmt%>" />
                <input type="hidden"  name="meldeart"   value="" />
                <input type="hidden"  name="kategorie"  value="" />
                <input type="hidden"  name="datum"      value="" />
                <input type="hidden"  name="filter"     value="" />
                <input type="hidden"  name="filterText" value="" />
                <input type="hidden"  name="spalte"     value="" />
                <input type="hidden"  name="suche"      value="" />
                <input type="hidden"  name="suchText2"  value="" />
                <input type="hidden"  name="monat"      value="" />
                <input type="hidden"  name="jahr"       value="" />
                <input type="hidden"  name="zeit"       value="" />
                <input type="hidden"  name="von"        value="" />
                <input type="hidden"  name="bis"        value="" />
                <input type="hidden"  name="intMaxRecord" value="" />
                <input type="hidden"  name="bTopDown"   value="" />
            </form>
        <%

        ' Connectionstring zusammenbauen #$DatabaseCall
        Dim conString
        Dim dbq
        dbq = historyPath & "\" & aktDatum & ".mdb"
        'conString= "Driver={Microsoft Access Driver (*.mdb)}; Dbq=" & dbq & "; Uid=; Pwd=;"
        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& dbq &";User Id=admin;Password=;"
        
        ' Verbindung herstellen
        Connection.Open conString
        Dim datum3,dummyDate1,dummyDate2
        Dim SQLStmtFirst1, SQLCount, SQLStmtNext1, SQLStmtPrev1, SQLStmt2, SQLStmt2Prev, SQLStmt
        if Err.Number = 0 then 'Wenn Datenbankverbindung ohne Fehler
            On Error Goto 0
            
            ' ----------------------------------------------------------------------------------------------------------------------------------------------------------
            SQLStmtFirst1 = "SELECT Top " & anzahlDatensaetze & " lfdNo, Liegenschaft, " & userAdr_SpaltenNameH & ", Plant, Messagetext, TStamp, Ereignistext, Nutzeradresse, PhyAdresse, Protokoll, Priority, Status FROM Prot WHERE (INStr('" & Session("AlarmLiegenschaftSet") & "',Liegenschaft & ';') > 0 OR INStr('WEBVISION',Liegenschaft) > 0) AND Typ=" & meldeKat2 & ""
            SQLCount = "SELECT * FROM Prot WHERE Typ=" & meldeKat2 & " AND (INStr('" & Session("AlarmLiegenschaftSet") & "',Liegenschaft & ';') > 0 OR INStr('WEBVISION',Liegenschaft) > 0)"
            if Not filterAktiv then 'Wenn Filter nicht aktiv
                If textSucheAktiv then
                    SQLStmtFirst1 = SQLStmtFirst1 & " AND " & spalte & " like '%" & suchText & "%'"
                    SQLCount = SQLCount & " AND " & spalte & " like '%" & suchText & "%'"
                End if
            else ' Wenn Filter aktiv
                if spalte="TStamp" then ' Datumsformat anpassen wenn auf Datumsspalte gesucht wird
                    datum3 = CDate(filterText)
                    datum3 = month(datum3) & "/" & day(datum3) & "/" & year(datum3) & " " & FormatDateTime(datum3, 3)
                    SQLStmtFirst1 = SQLStmtFirst1 & " AND " & spalte & "=#" & datum3 & "#"
                    SQLCount = SQLCount & " AND " & spalte & "=#" & datum3 & "#"
                else
                    If UCase(spalte) = "PRIORITY" then
                        SQLCount = SQLCount & " AND " & spalte & "=" & filterText & ""
                        SQLStmtFirst1 = SQLStmtFirst1 & " AND " & spalte & "=" & filterText & ""
                    Else
                        SQLCount = SQLCount & " AND " & spalte & "='" & filterText & "'"
                        SQLStmtFirst1 = SQLStmtFirst1 & " AND " & spalte & "='" & filterText & "'"
                    End If
                end if
            end if
            SQLStmtNext1  = SQLStmtFirst1 & " AND lfdNo < "
            SQLStmtPrev1  = SQLStmtFirst1 & " AND lfdNo > "
            
            If sortOrder = "" then ' Default sort order
                SQLStmt2 = " ORDER BY lfdNo DESC"
                SQLStmt2Prev = " ORDER BY lfdNo ASC"
            Else ' User defined sort order
                If instr(sortOrder, "DESC") > 0 then
                    SQLStmt2 = " ORDER BY " & sortOrder
                    SQLStmt2Prev = " ORDER BY " & replace(sortOrder, "DESC", "ASC")
                Else ' sortOrder ASC
                    SQLStmt2 = " ORDER BY " & sortOrder
                    SQLStmt2Prev = " ORDER BY " & replace(sortOrder, "ASC", "DESC")
                End If
            End If


        ' ----------------------------------- [NAVIGATION] --------------------------------------------
            Dim strPageDirection, nLastIndex
            strPageDirection = Trim(Request.QueryString("Page"))
            If "" = strPageDirection Then strPageDirection = "Start"
            nLastIndex = Trim(Request.QueryString("Index"))
            If ("" = nLastIndex Or Not IsNumeric(nLastIndex)) Then strPageDirection = "Start"

            Select Case strPageDirection
                Case "Next"
                    SQLStmt = SQLStmtNext1 & nLastIndex & SQLStmt2
                Case "Prev"
                    SQLStmt = SQLStmtPrev1 & nLastIndex & SQLStmt2Prev
                    bTopDown = True
                Case Else
                    SQLStmt = SQLStmtFirst1 & SQLStmt2
            End Select
            ' Debug-Ausgabe --> Bei Bedarf ausgeben...
            ' Response.Write SQLStmt
            ' Response.End
            

            %>
            <!-- Form zum Export aller aufgelisteten Alarme (Parameterübergabe nur per "Post" möglich, da "Get" zu wenig Zeichen zulässt (Url-Parameter beschränkt auf 2000 Zeichen)) -->
            <form name="ExportForm" method="post" action="alarmExport.asp" target="_self">
                <input type="hidden"  name="sql"        value="<%=SQLStmt%>" />
                <input type="hidden"  name="meldeart"   value="" />
                <input type="hidden"  name="kategorie"  value="" />
                <input type="hidden"  name="datum"      value="" />
                <input type="hidden"  name="filter"     value="" />
                <input type="hidden"  name="filterText" value="" />
                <input type="hidden"  name="spalte"     value="" />
                <input type="hidden"  name="suche"      value="" />
                <input type="hidden"  name="suchText2"  value="" />
                <input type="hidden"  name="monat"      value="" />
                <input type="hidden"  name="jahr"       value="" />
                <input type="hidden"  name="zeit"       value="" />
                <input type="hidden"  name="von"        value="" />
                <input type="hidden"  name="bis"        value="" />
                <input type="hidden"  name="intMaxRecord" value="" />
                <input type="hidden"  name="bTopDown"   value="" />
            </form>
            <%
            '--------------------------------------- [AUTOFILTER] --------------------------------------------
            if autoFilter then
            %>
                <!--#include file="autoFilterIncludeH.asp"-->
            <%
            End if
            '--------------------------------------- [/AUTOFILTER] --------------------------------------------


            Dim bNoRecords, rs2
            Set rs2 = CreateObject ("ADODB.Recordset")
            rs2.Cursortype = adOpenKeyset
            rs2.Locktype = adLockOptimistic
            rs2.Open SQLStmt, Connection

            If Not rs2.EOF and Not rs2.BOF Then
                avarRecords = rs2.GetRows()
            Else
                bNoRecords = True
            End If

            'Maximal Anzahl an Datensätzen ermitteln...
            intMaxRecord = getMaxCount(SQLCount)
            rs2.Close
            '------------- /Maximale Anzahl Datensätze -------------
            
            Dim nPrevPaging,nLastPaging
            Set rs2 = Nothing
            If Not bNoRecords Then
                nRecords = UBound(avarRecords,2)
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
            tdcolor="EEEEEE"
            
            If Not bTopDown Then
                For intRecord = 0 To nRecords
                    call WriteRecord(avarRecords, intRecord)
                    if Err.Number = 0 then
                        currentRecord=currentRecord+1
                    End if
                    Response.Flush
                Next
            Else
                For intRecord = nRecords To 0 Step -1
                    call WriteRecord(avarRecords, intRecord)
                    if Err.Number = 0 then
                        currentRecord=currentRecord+1
                    End if
                    Response.Flush
                Next
            End If

            Dim maxCount
            maxCount=0
           
        ' ----------------------------------- [/NAVIGATION] --------------------------------------------
            If recordAnzCount = 0 then
                Response.Write "</tr>"
            End If
        End If
        Err.Clear
        '--------------------------------- [TAGESANSICHT] -------------------------------------

    '============================================== [ /HISTORISCHE ALARME ] ==============================================


    '========================================= [ AKTUELLE / QUITTIERTE ALARME ] ==============================================

    else 'Alle anderen Alarme (aktive/quittierte)

        '--------------------------------- [CONNECTION AUFBAUEN] -------------------------------------
        Connection.Open Session("ConnectionString")
        Error.clear
        'On Error Resume Next
        If trim(sortOrder) = "" then
            If meldeArt2 = "255" then 'gesperrte Meldungen
                sortOrder = " ORDER BY LogTime DESC"
            Else
                sortOrder = " ORDER BY DT_Stamp_high DESC"
            End If
        Else
            sortOrder = " ORDER BY " & sortOrder
        End If

        SQLStmt = "SELECT Meldeart, "& userAdr_SpaltenName &", UserAdr, Lieg, DT_Stamp_high, PlantID, Alarmtext, " & _
                  " Path, Ust, Reg, quittime, quit_user, prio, LogTime, Sondertextdatei, dv_path, Protokoll, " & _
                  " PhyAdresse, AckedTrans, BACnetStatus, quitted  FROM Sondertexte" & _
                  " WHERE Meldeart=" &CInt(meldeKat2)& ""
        if Not filterAktiv then 'Wenn Filter nicht aktiv
            If textSucheAktiv then
                If UCase(spalte) = "PRIO" then
                    If meldeArt2 = "255" then 'gesperrte Meldungen
                        SQLStmt = SQLStmt & " AND "&spalte&" = "&suchText&" AND prio BETWEEN 250 AND " & _
                            CInt(meldeArt2) &" AND trim(alarmtext) <> '_'"
                    Else
                        SQLStmt = SQLStmt & " AND "&spalte&" = "&suchText&" AND quitted="& CInt(meldeArt2) & _
                            " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND trim(alarmtext)<>'_'"
                        'Response.Write "SQLStmt: " & SQLStmt
                    End If
                Else
                    If meldeArt2 = "255" then 'gesperrte Meldungen
                        SQLStmt = SQLStmt & " AND "&spalte&" like '%"&suchText&"%' AND prio BETWEEN 250 AND " & _
                            CInt(meldeArt2) &" AND trim(alarmtext) <> '_'"
                    Else ' All active alarms
                        If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                            SQLStmt = SQLStmt & " AND "&spalte&" like '%" & suchText & "%' AND high_low=-1 AND prio " & _
                                "BETWEEN 1 AND 250 AND trim(alarmtext) <> '_'"
                        Else
                            SQLStmt = SQLStmt & " AND "&spalte&" like '%"&suchText&"%' AND quitted="& CInt(meldeArt2) & _
                                " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND trim(alarmtext) <> '_'"
                        End If
                    End If
                End If
            Else
                If meldeArt2 = "255" then 'gesperrte Meldungen
                    SQLStmt = SQLStmt & " AND prio BETWEEN 250 AND "& CInt(meldeArt2) & _
                        " AND trim(alarmtext) <> '_' "
                Else
                    If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                        SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND trim(alarmtext) <> '_' "
                    Else
                        SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND prio BETWEEN 1 " & _
                            "AND 250 AND trim(alarmtext) <> '_'"
                    End If
                End If
            End If
        else ' Wenn Filter aktiv
            If meldeArt2 = "255" then 'gesperrte Meldungen
                SQLStmt = SQLStmt & " AND "&spalte&"='"&filterText&"' AND prio BETWEEN 250 AND "& CInt(meldeArt2) & _
                    " AND trim(alarmtext) <> '_'"
            Else
                If UCase(spalte) = "DT_STAMP_HIGH" then 'Bei Datumsfilter, Datum und SQL-Statement anpassen.
                    Dim datum4
                    datum4 = CDate(filterText)
                    datum4 = month(datum4)&"/"&day(datum4)&"/"&year(datum4)&" "& FormatDateTime(datum4,3)
                    If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                        SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250" & _
                            " AND trim(alarmtext) <> '_' AND " & spalte & "=#" & datum4 & "#"
                    Else
                        SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND " & _
                            "prio BETWEEN 1 AND 250 AND trim(alarmtext) <> '_' AND " & spalte & "=#" & datum4 & "#"
                    End If
                Else
                    If UCase(spalte) = "PRIO" then ' If selected autofilter column is "prio"
                        If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                            SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND" & _
                            " trim(alarmtext) <> '_' AND " & spalte & "=" & filterText
                        Else
                            SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND " & _
                                "prio BETWEEN 1 AND 250 AND trim(alarmtext) <> '_' AND " & spalte & "=" & filterText
                        End If
                    Else
                        If CInt(meldeArt2) = -255 then ' Alle aktiven Alarme (auch die Quittierten)
                            SQLStmt = SQLStmt & " AND high_low=-1 AND prio BETWEEN 1 AND 250 AND" & _
                            " trim(alarmtext) <> '_' AND " & spalte & "='" & filterText & "'"
                        Else
                            SQLStmt = SQLStmt & " AND quitted="& CInt(meldeArt2) &" AND high_low=-1 AND " & _
                                "prio BETWEEN 1 AND 250 AND trim(alarmtext) <> '_' AND " & spalte & "='" & filterText & "'"
                        End If
                    End If
                End If
            End If
        end if
        SQLStmt = SQLStmt & sortOrder
        ' Add sort order to statement
        
        '--------------------------------- [/CONNECTION AUFBAUEN] -------------------------------------

        'Response.Write SQLStmt
        '--------------------------------------- [AUTOFILTER] --------------------------------------------
        If autoFilter then
        %>
            <!--#include file="autoFilterIncludeA.asp"-->
        <%
        End if
        '--------------------------------------- [/AUTOFILTER] --------------------------------------------


        '-------------------------------------- [TABELLE FÜLLEN]-------------------------------------------
        'Response.Write SQLStmt
        'On Error GoTo 0
        Set RS = Connection.Execute(SQLStmt)
        tdcolor="bgRowColor3"
        Dim dProtokoll
        Dim formContainer : formContainer = ""
        Dim fBackValue : fBackValue = ""

        Do While CheckRS(RS)
            dProtokoll = ""
            dProtokoll = RS("Protokoll").Value

            if instr(Session("AlarmLiegenschaftSet"),CheckRequest(RS, "Lieg") & ";") > 0 OR instr("WEBVISION",CheckRequest(RS, "Lieg")) > 0 then
                if tdcolor="bgRowColor3" then
                    tdcolor="bgRowColor2"
                else
                    tdcolor="bgRowColor3"
                end if

                ' Put the form-html-code into a variable and print them at the end of the table on the side.
                fBackValue = Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING")
                formContainer = formContainer & "<form method=""post"" name=""F_" & currentRecord & """ action=""$Quittung.asp"">"        
                formContainer = formContainer & "<input type=""hidden"" name=""Path"" value=""" & CheckRequest(RS, "Path") & """ />"
                formContainer = formContainer & "<input type=""hidden"" name=""UST"" value=""" & CheckRequest(RS, "UST") & """ />"
                formContainer = formContainer & "<input type=""hidden"" name=""Reg"" value=""" & CheckRequest(RS, "Reg") & """ />"
                formContainer = formContainer & "<input type=""hidden"" name=""SQL"" value=""" & SQLStmt & """ />"
                formContainer = formContainer & "<input type=""hidden"" name=""fBack"" value=""" & fBackValue & """ />"
                formContainer = formContainer & "</form>"
                
                ' If event is quitted and event view is "all active and active quitted events" then
                ' color the font to "gold".
                Dim quitAllowed
                quitAllowed = false
                If RS("quitted").Value = -1 AND (meldeArt2 = "-255" OR meldeArt2 = "-1") then
                    If meldeArt2 = "-255" then
                        tdcolor = "bgQuittedColor"
                        quitAllowed = false
                    Else
                        quitAllowed = true
                    End If
                Else
                    quitAllowed = true
                End If
                %>
                <tr class="<%=tdcolor%>" id="Zeile<%=currentRecord%>" onclick="setRowColor('<%=tdcolor%>',<%=currentRecord%>)" ondblclick="IWin('<% = Replace(Replace(Replace(Replace(CheckRequest(RS,"Path"),"\","~")," ","*"),"&","{"),"+","$$43")%>',<% = CheckRequest(RS,"UST")%>,<% = CheckRequest(RS,"REG")%>,<%=dProtokoll%>)">
                    <td style="width:16px;">
                        <img src="<%=getEventImage(RS("quitted").Value, uStyle)%>" title="<%=getEventText(meldeKat2)%>" alt="<%=getEventText(RS("Meldeart").Value)%>" />
                    </td>

                    <td valign="middle" class="<%=tableCellClassName %>">
                    <%
                    ' Abfragen ob gesperrte Alarme angezeigt werden
                    If meldeArt2 = "255" then ' Wenn ja, LogTime-Spalte anzeigen...
                        Response.Write CheckRequest(RS, "LogTime")
                    Else ' ...ansonsten normale Datumsspalte anzeigen
                        Response.Write CheckRequest(RS, "DT_Stamp_high")
                    End If
                    %>
                     </td>
                    
                    <td valign="middle" class="<%=tableCellClassName %>">
                        <%= replace(CheckRequest(RS, "Alarmtext"),"_"," ") %>
                    </td>
                    <td valign="middle" class="<%=tableCellClassName %>">
                        <%= replace(CheckRequest(RS, "Lieg"),"_"," ") %>
                    </td>
                    <td valign="middle" class="<%=tableCellClassName %>">
                        <%= replace(CheckRequest(RS, "PlantID"),"_"," ") %>
                    </td>
                    <td valign="middle" class="<%=tableCellClassName %>">
                        <%= CheckRequest(RS, "prio") & "&nbsp;(" & arrPriority(CInt(RS("prio").value)) & ")"%>
                    </td>
                    <%
                    Dim linkColor : linkColor = ""
                    If meldeArt2 = "0" Or meldeArt2 = "-255" then
                        linkColor = "red" 
                    End If
                    %>
                    <% if Session("optUseradr") = 1 then %>
                            <td valign="middle"  class="<%=tableCellClassName %>">
                            <% If CheckRequest(RS, "Reg") = "-1" then 'If event is a network error message (no regular data point object) %>
                                <a style="color:<%=linkColor%>; text-decoration:underline;" href="javascript:IWin('<% = Replace(Replace(Replace(Replace(CheckRequest(RS,"Path"),"\","~")," ","*"),"&","{"),"+","$$43")%>',<% = CheckRequest(RS,"UST")%>,<% = CheckRequest(RS,"REG")%>,<% = dProtokoll %>)">
                            <% Else %>
                                <a style="color:<%=linkColor%>; text-decoration:underline;" href="javascript:details('adr2','/webvisionnt/zsystem/maindata/detail.asp?useradr=<%=replace(RS("UserAdr"),chr(34),"")%>','920','665')">
                            <% End If%>
                                <%= replace(CheckRequest(RS, userAdr_SpaltenName),"_"," ") %>
                            </a>
                        </td>
                    <% end if %>

                    <% If qUserRight AND meldeArt2 = "-1" then %>
                        <td valign="middle"  class="<%=tableCellClassName %>">
                            <%=CheckRequest(RS, "Quit_User") %>
                        </td>
                    <% End If %>
                    
                    <% if meldeArt2 = "0" OR meldeArt2 = "-255" then %>
                        <!-- (Active alarms) or (active and active acknowledged alarms) -->
                        <td>
                            <table style="margin-left:5px; margin-top:0px;" cellpadding="0" cellspacing="0">
                                <tr>
                            <% if Session("CurLevel") > 2 then
                                dummyVal = CheckRequest(RS, "UST") & CheckRequest(RS, "Reg")
                                If RS("Protokoll").Value = 1 then 'BACnet-Event/Alarm
                                    ' Thinking about that, if this is the right way!
                                    ' Normaly the BACnet-bits are in the right state and all
                                    ' acknowledge-buttons are correct.
                                    If quitAllowed then
                                        'Quittierung
                                        Dim S1, S2, S3
                                        Dim quittBits
                                        quittBits = CheckRequest(RS, "AckedTrans")

                                        If quittBits <> "" then
                                            S1=mid(quittBits,1,1)
                                            S2=mid(quittBits,2,1)
                                            S3=mid(quittBits,3,1)
                                        Else
                                            S1="T"
                                            S2="T"
                                            S3="T"
                                        End If

                                        Dim benQ1, benQ2, benQ3

                                        If S1="F" then benQ1=ben150 Else benQ1=ben151 End If
                                        If S2="F" then benQ2=ben150 Else benQ2=ben151 End If
                                        If S3="F" then benQ3=ben150 Else benQ3=ben151 End If

                                        If S1="T" then ' Schon quittiert %>
                                            <td width="20">
                                            <a href="#"><img border="0" src="../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_<%=S1%>.gif" alt="<%=benQ1%>:&nbsp;<%=ben152%>" title="<%=benQ1%>:&nbsp;<%=ben152%>" /></a>
                                            </td>
                                        <% Else %>
                                            <td width="20">
                                            <a href="javascript:quittalarm('<%=RS("PhyAdresse")%>~2','<%=CheckRequest(RS, "Lieg")%>', '<%=CheckRequest(RS, userAdr_SpaltenName)%>', '<%= RS("Alarmtext").value %>')"><img border="0" src="../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_<%=S1%>.gif" alt="<%=benQ1%>:&nbsp;<%=ben152%>" title="<%=benQ1%>:&nbsp;<%=ben152%>" /></a>
                                            </td>
                                        <% End If %>
                                        <% If S2="T" then ' Schon quittiert %>
                                            <td width="20">
                                            <a href="#"><img border="0" src="../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_<%=S2%>.gif" alt="<%=benQ2%>:&nbsp;<%=ben153%>" title="<%=benQ2%>:&nbsp;<%=ben153%>" /></a>
                                            </td>
                                        <% Else %>
                                            <td width="20">
                                            <a href="javascript:quittalarm('<%=RS("PhyAdresse")%>~1','<%=CheckRequest(RS, "Lieg")%>', '<%=CheckRequest(RS, userAdr_SpaltenName)%>', '<%= RS("Alarmtext").value %>')"><img border="0" src="../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_<%=S2%>.gif" alt="<%=benQ2%>:&nbsp;<%=ben153%>" title="<%=benQ2%>:&nbsp;<%=ben153%>" /></a>
                                            </td>
                                        <% End If %>
                                        <% If S3="T" then ' Schon quittiert %>
                                            <td width="20">
                                            <a href="#"><img border="0" src="../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_<%=S3%>.gif" alt="<%=benQ3%>:&nbsp;<%=ben154%>" title="<%=benQ3%>:&nbsp;<%=ben154%>" /></a>
                                            </td>
                                        <% Else %>
                                            <td width="20">
                                            <a href="javascript:quittalarm('<%=RS("PhyAdresse")%>~0','<%=CheckRequest(RS, "Lieg")%>', '<%=CheckRequest(RS, userAdr_SpaltenName)%>'), '<%= RS("Alarmtext").value %>'"><img border="0" src="../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_<%=S3%>.gif" alt="<%=benQ3%>:&nbsp;<%=ben154%>" title="<%=benQ3%>:&nbsp;<%=ben154%>" /></a>
                                            </td>
                                        <% End If
                                    Else
                                        %>
                                        <td colspan="3"><%=ben151 & " (" & CheckRequest(RS, "Quit_User") & ")"%></td>
                                        <%
                                    End If
                                Else
                                    If quitAllowed then
                                        %>
                                        <td width="20">
                                        <a href="javascript:window.document.F_<%=currentRecord%>.submit(); logAcknowledge('<%=RS("Alarmtext").value %>', '<%=RS("Lieg").value %>', '<%=RS(userAdr_SpaltenName).value %>')" class="textLink">
                                            <img style="vertical-align:text-top; margin-right:4px;" 
                                                src="../styles/<%=uStyle%>/bin/bacAlarm/quitt_0_F.gif" border="0" 
                                                id="Image1" name="Image1" alt="<%=ben150%>:&nbsp;<%=ben152%>" title="<%=ben150%>:&nbsp;<%=ben152%>" />
                                        </a>
                                        </td>
                                        <td width="20">&nbsp;</td>
                                        <td width="20">&nbsp;</td>
                                        <%
                                    Else
                                        %>
                                        <td colspan="3"><%=ben151 & " (" & CheckRequest(RS, "Quit_User") & ")"%></td>
                                        <%
                                    End If
                                End If
                                %>
                            <% end if %>
                            </tr>
                            </table> 
                        </td>
                    <% End if %>
                    <%
                    Dim colspan2
                    colspan2 = 3
                    If meldeArt2 = "0" Then  'Aktive Alarme
                        colspan2=2
                    End If
                    %>
                    <td align="left" colspan="<%=colspan2%>">
                        <%
                        ' Abfragen ob eine Zusatzinfo verfügbar ist,...
                        Dim bildNameGrafik, bildNameText, bildNameInfo, std, dVP
                        std = CheckRequest(RS, "Sondertextdatei")
                        dVP = CheckRequest(RS, "dv_path")
                        bildNameInfo = "<img src=""../styles/" & uStyle & "/bin/info_help.gif"" title=""" & ben088 & """ alt=""" & ben088 & """ style=""margin-right:4px;"" border=""0"">"
                        
                        ' Zusatztext vorhanden?
                        If std = "" OR std = "LEER" OR std = "-/-" OR isNull(std) then
                            bildNameText = ""
                        Else '...wenn ja, animierten Info-Button einblenden
			                Dim slash
			                slash=""
			                ' Überprüfen, ob erstes Zeichen ein Slash ist
			                If left(std,1) <> "/" then
				                slash="/" ' Wenn nicht, dann Slash-Variable füllen...
			                Else
				                slash=""  ' ...ansonsten leeren.
			                End If

			                ' Wenn nur ein Zusatztext eingegeben und keine Datei ausgewählt wurde
			                ' Zusatztext anzeigen
			                ' Gekennzeichnet durch das Zeichenkürzel "|ZT|"
			                If left(std,4) = "|ZT|" then
				                std=mid(std,5,Len(std))
                                bildNameText = "<img src=""../styles/" & uStyle & "/bin/txt.gif"" title=""" & std & """ alt=""" & std & """ style="" margin-right:4px; cursor:help;"" border=""0"">"
			                else
					            ' Pfad relativieren
					            num = InStr(Session("Projektpfad"),"WEBVISIONNT")
					            pfad = "/"&Replace(mid(Session("Projektpfad"),num,Len(Session("Projektpfad"))),"\","/")
                                pfad = pfad &"/Alarm/Zusatztexte"& slash & std                    
                                bildNameText = "<a href="""  & pfad & """ target=""_blank"">"
                                bildNameText = bildNameText & "<img src=""../styles/" & uStyle & "/bin/txt.gif"" title=""" & ben182 & """ alt=""" & ben182 & """ style="" margin-right:4px;"" border=""0"">"
                                bildNameText = bildNameText & "</a>"
                            End If
                            
                        End If
                        
                        ' Zusatzgrafik vorhanden?
                        If dVP = "" OR dVP = "LEER" OR dVP = "-/-" OR isNull(dVP) then
                            'bildName="../styles/"& uStyle&"/bin/info_help.gif"
                            bildNameGrafik = ""
                        Else
                            'bildName="../styles/"& uStyle&"/bin/menu/tree_anlage.gif"
                            bildNameGrafik = "<a href=""" & dVP & """ target=""_blank"">"
                            bildNameGrafik = bildNameGrafik & "<img src=""../styles/" & uStyle & "/bin/menu/tree_anlage.gif"" title=""" & ben181 & """ alt=""" & ben181 & """ style=""margin-right:4px;"" border=""0"">"
                            bildNameGrafik = bildNameGrafik & "</a>"
                        End If
                        std=""
                        dVP=""
                        %>
                        <table style="margin-left:5px; margin-top:0px;" cellpadding="0" cellspacing="0">
                            <tr>
                                <td width="20"><a href="javascript:IWin('<% = Replace(Replace(Replace(Replace(CheckRequest(RS,"Path"),"\","~")," ","*"),"&","{"),"+","$$43")%>',<% = CheckRequest(RS,"UST")%>,<% = CheckRequest(RS,"REG")%>,<% = dProtokoll %>)"><%= bildNameInfo%></a></td>
                                <td width="20">
                                    <a href="javaScript:openDetails('<%= RS("Useradr").value %>');"><img alt="<%=ben179 %>" title="<%=ben179 %>" src="../styles/<%=ustyle%>/bin/alarme/dpProperties.gif" style="margin-right:4px;" border="0" /></a>
                                </td>
                                <!--<td width="20"><a href="javaScript:openMeldeManager('<%= RS("Useradr").value %>');">
                                    <img alt="<%=ben063 %>" title="<%=ben063 %>" src="../styles/<%=ustyle%>/bin/alarme/meldekonf.gif" 
                                        style="margin-right:4px;" border="0" />
                                    </a>
                                </td>-->
                                <td width="20"><%= bildNameGrafik%></td>
                                <td width="20"><%= bildNameText%></td>
                            </tr>
                        </table>
                        
                    </td>
                </tr>

            
            <%
            currentRecord=currentRecord+1
            end if
            RS.MoveNext
            Response.Flush
        Loop
        '-------------------------------------------- [/TABELLE FÜLLEN]-------------------------------------------
    End if
    '========================================= [ /AKTUELLE - QUITTIERTE ALARME ] ==============================================


    '------------------------ [NAVIGATION HISTORISCHE ALARME] --------------------------------
    If meldeArt2 = "1" then
        Dim letzteDatensaetze
        letzteDatensaetze=0
        ' Bei Vorwärts
        If Trim(Request.QueryString("Page")) = "Next" then
            letzteDatensaetze = CInt(Request.QueryString("letzteDatensaetze")) + recordAnzCount
        ' Bei Rückwärts
        elseif Trim(Request.QueryString("Page")) = "Prev" then
            letzteDatensaetze = CInt(Request.QueryString("letzteDatensaetze")) - CInt(Request.QueryString("lastRecord"))
        ' Am Anfang
        else
            ' Wenn die Seite zum ersten mal aufgerufen wird
            If recordAnzCount = "" then
                letzteDatensaetze = 0 ' Count auf 0 stellen
            ' Wenn beim Zurückgehen auf den Anfang gesprungen wird
            Else
                letzteDatensaetze = recordAnzCount ' Count auf übergebenen Wert stellen
            End If
        End If
        'Response.Write "Anzahl DATENSÄTZE: "&recordAnzCount
        %>

        <tr bgcolor="white">
            <td colspan="<%=colspan%>">
                <% If intMaxRecord = 1 then %>
                    <%=ben028%>&nbsp;<%=letzteDatensaetze-recordAnzCount&"&nbsp;-&nbsp;"&letzteDatensaetze%>&nbsp;&nbsp;<%=ben031%>&nbsp;<%=intMaxRecord%>&nbsp;<%=ben028%>
                <% Else %>
                    <%=ben029%>&nbsp;<%=letzteDatensaetze-recordAnzCount&"&nbsp;-&nbsp;"&letzteDatensaetze%>&nbsp;&nbsp;<%=ben031%>&nbsp;<%=intMaxRecord%>&nbsp;<%=ben030%>
                <% End If%>
                <img src="../styles/<%=uStyle%>/bin/alarme/prev.gif" border="0" alt="<%=ben054%>" style="vertical-align:bottom; cursor:hand; position:absolute; right:35px;" onclick="goWay('Prev')">
                <img src="../styles/<%=uStyle%>/bin/alarme/next.gif" border="0" alt="<%=ben055%>" style="vertical-align:bottom; cursor:hand; position:absolute; right:10px;" onclick="goWay('Next')">
            </td>
        </tr>
        </table>
        <%
    Else
    %>
        </table>
    <%
    End If
    '------------------------ [/NAVIGATION HISTORISCHE ALARME]--------------------------------

Response.Write formContainer
' Wenn keine Meldung vorhanden entsprechende Kenntlichmachung
If currentRecord = 0 Then
%>

    <% if meldeArt2 = "1" then %>
        <div class="korbLeer hist">
            <%=ben130%>
        </div>
    <%  Else %>
        <div class="korbLeer">
            <%=ben130%>
        </div>
    <%  End if %>

<%  End if %>

<script language="javascript" type="text/javascript">

    /* Aktualisierungsframe (treehaeder) Kategorie mitgeben -- Wenn Historische Meldungen, keine Aktualisierung */
    /* Wird beim Starten der Seite aufgerufen (<body onload-Ereignis>) */
    var meldeArt = null;
    function setHistory(wert, art) 
    {
        //window.parent.treeheader.window.location.href = "aktFrame.asp?meldeKat=" + wert + "";
        window.top.focus();
        setFilterAndTextSearch();
        meldeArt = wert;
    }

    function setFilterAndTextSearch()
    {
        try
        {
        <% If NOT autoFilter then %>
            if(window.top.Oben)
            {
                if(window.top.Oben.document.getElementById("btnAutofilter").className == "button3Active")
                {
                    window.top.Oben.document.getElementById("btnAutofilter").className = "button3";
                    window.top.Oben.document.getElementById("AutoFilter").src = "../styles/<%=Session("userStyle")%>/bin/autofilter.gif";
                }
            }
        <% End If %>
        <% If NOT textSuche then %>
            if(window.top.Oben)
            {
                if(window.top.Oben.document.getElementById("btnTextSuche").className == "button3Active")
                {
                    window.top.Oben.document.getElementById("btnTextSuche").className = "button3";
                    window.top.Oben.document.getElementById("TextSuche").src = "../styles/<%=Session("userStyle")%>/bin/suchen.gif";
                }
            }
        <% End If %>
        }
        catch(e)
        {}
    }

    var recordPresent;
    recordPresent = <%=currentRecord%>;
    var http = null;
    /* ------------------------------------------------------------------------------------------------- */
    /* Quittieren */
    function quittalarm(qtxt, lieg, userAdr, alarmText)
    {
        //alert(qtxt + " - " + lieg + " - " + userAdr);
        /* XMLHttp - Objekt erzeugen (Abfrage auf Microsoft und andere)*/
        if (window.XMLHttpRequest)
        {
            http = new XMLHttpRequest();
        }
        else if (window.ActiveXObject)
        {
            http = new ActiveXObject("Microsoft.XMLHTTP");
        }

        /* Wenn Objekt erfolgreich erzeugt wurde */
        if (http != null)
        {
            http.open("POST", "quittAlarme_ajax.asp", true); // Seite aufrufen
            http.onreadystatechange = function()
            {
                if (http.readyState == 4)
                {
                    /* Log this user action */
                    logAcknowledge(alarmText, lieg, userAdr);
                }
            };
            http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");   // header setzen
            //http.send("ua="+ ChangeUmlInhtml2(userAdr) +""); // Parameter an die Seite übergeben
            http.send("q="+ ChangeUmlInhtml2(qtxt) +"&l="+ ChangeUmlInhtml2(lieg) +"&u="+ ChangeUmlInhtml2(userAdr) +"");  // Parameter an die Seite übergeben
            // Log user action
            //logAcknowledge(qtxt,lieg,userAdr);
        }
    }
    /* ------------------------------------------------------------------------------------------------- */


    /**
     * Log the user action into the event db
     */
    function logAcknowledge(messageText, lieg, userAdr)
    {
        if(http == null)
        {
            /* XMLHttp - Objekt erzeugen (Abfrage auf Microsoft und andere)*/
            if (window.XMLHttpRequest)
            {
                http = new XMLHttpRequest();
            }
            else if (window.ActiveXObject)
            {
                http = new ActiveXObject("Microsoft.XMLHTTP");
            }
        }

        /* Wenn Objekt erfolgreich erzeugt wurde */
        if (http != null)
        {
            http.open("POST", "/webvisionnt/zsystem/maindata/eventProt.asp", true); // Seite aufrufen
            http.onreadystatechange = function(){if (http.readyState == 4){/*Do nothing*/}};
            http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");   // header setzen
            http.send("q="+ ChangeUmlInhtml2(messageText) +"&l="+ ChangeUmlInhtml2(lieg) +"&u="+ ChangeUmlInhtml2(userAdr) +"");  // Parameter an die Seite übergeben
        }
    }

    /* Seite mit neuer Anzahl an anzuzeigenden Zeilen neu laden. */
    function loadNewMaxRows()
    {
        maxRows= window.document.getElementById("historylength").value;
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=<%=filter%>&filterText=<%=filterText%>&suche=<%=suche%>&suchText2=<%=ChangeUmlInhtml2(suchText)%>&spalte=<%=suchSpalte%>&newMaxRow="+maxRows+"";
    }

    /* Durch die Ergebnisse steppen */
    function goWay(direction)
    {
        if(direction=="Prev") // <-- Zurück
        {
            <% If letzteDatensaetze-recordAnzCount > 0 then %>
                window.location.href = "alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=<%=filter%>&filterText=<%=filterText%>&suche=<%=suche%>&suchText2=<%=ChangeUmlInhtml2(suchText)%>&spalte=<%=suchSpalte%>&Page=Prev&Index=<%=nPrevPaging%>&letzteDatensaetze=<%=letzteDatensaetze%>&lastRecord=<%=recordAnzCount%>&monat=<%=aktMonat%>&jahr=<%=aktJahr%>";
            <% End If %>
        }
        else // --> Vor
        {
            <% ="var d='"&letzteDatensaetze&"';" %>
            <% If letzteDatensaetze < intMaxRecord then %>
                window.location.href = "alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=<%=filter%>&filterText=<%=filterText%>&suche=<%=suche%>&suchText2=<%=ChangeUmlInhtml2(suchText)%>&spalte=<%=suchSpalte%>&Page=Next&Index=<%=nLastPaging%>&letzteDatensaetze=<%=letzteDatensaetze%>&monat=<%=aktMonat%>&jahr=<%=aktJahr%>";
            <% End If %>
        }
    }

    /* Jahresansicht wechseln */
    function changeYear(year)
    {
        window.document.all.jahrSelect.disabled=true;
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&datum=-&kategorie=<%=meldeKat2%>&monat=|&jahr="+year+"&ansicht=<%=ansicht%>&filter=<%=filter%>&suche=<%=suche%>";
    }

    /* Monatsansicht wechseln */
    function changeMonth(month)
    {
        window.document.all.monatSelect.disabled=true;
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&datum=-&kategorie=<%=meldeKat2%>&monat="+month+"&jahr=<%=aktJahr%>&ansicht=<%=ansicht%>&filter=<%=filter%>&suche=<%=suche%>";
    }

    /* Tagesansicht wechseln */
    function changeDay(day)
    {
        window.document.all.zeit.disabled=true;
        window.location.href="historyMonthWait.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum="+day+"&monat=<%=aktMonat%>&jahr=<%=aktJahr%>&filter=<%=filter%>&suche=<%=suche%>&filterText=<%=ChangeUmlInhtml2(filterText)%>&spalte=<%=spalte%>";
    }

    /* Info´s zur historischen Alarmmeldung anzeigen */
    function historyInfo(lfdNo, zIndex, protokoll)
    {
        window.open("historyInfo.asp?lfdNo=" + lfdNo + "&aktDatum=<%=aktDatum%>&kategorie=<%=meldeKat2%>&protokoll=" + 
            protokoll + "", "Info","height=550,width=600,top=200,left=200,status=no,toolbar=no,menubar=no,location=no");
    }

    /* Info´s zur aktuellen/quittierten Alarmmeldung anzeigen */
    function aktInfo(usAdr)
    {
        window.open("aktInfo.asp?userAdr="+usAdr+"&protokoll=<%=dpProtokoll%>", "Info","height=310,width=600,top=200,left=200,status=no,toolbar=no,menubar=no,location=no, resizable=yes");
    }

    /* Info´s zur Alarmmeldung anzeigen */
    function IWin(xPath,xUST,xReg)
    {
        window.open("aktInfo.asp?vPath="+xPath  +"&vUST="+xUST  +"&vReg="+xReg+"&kategorie=<%=meldeKat2%>&meldeart4=<%=meldeArt2%>",  "Info","height=330,width=600,top=200,left=200,status=no,toolbar=no,menubar=no,location=no ,resizable=yes");
    }

    /* Autofilter setzen */
    function setAutoFilter()
    {
        window.location.href="filterHWait.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=ja&monat=<%=aktMonat%>&jahr=<%=aktJahr%>&ansicht=<%=ansicht%>";
    }

    /* Autofilter loeschen */
    function clearAutoFilter()
    {
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=nein&monat=<%=aktMonat%>&jahr=<%=aktJahr%>&ansicht=<%=ansicht%>";
    }

    /* Gefilterte Seite laden --> Wenn Anwender einen Filter auswählt */
    function loadFilter(filter, spalte)
    {
        var adresse;
        var dumm = filter;
        filter = ChangeUmlInhtml2(dumm);
        adresse = "alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=ja&filterText="+filter+"&spalte="+spalte+"&monat=<%=aktMonat%>&jahr=<%=aktJahr%>&ansicht=<%=ansicht%>";
        window.location.href=adresse;
    }

    /* Textsuche setzen */
    function setTextSuche()
    {
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=nein&suche=ja&monat=<%=aktMonat%>&jahr=<%=aktJahr%>";
    }

    /* Textsuche loeschen */
    function clearTextSuche()
    {
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=nein&suche=nein&monat=<%=aktMonat%>&jahr=<%=aktJahr%>";
    }

    /* Textsuche laden --> Wenn Anwender eine Textsuche ausführt */
    function loadSuche(suchText, spalte)
    {
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=nein&suche=ja&suchText2="+suchText+"&spalte="+spalte+"&monat=<%=aktMonat%>&jahr=<%=aktJahr%>";
    }

    /* Ansicht wechseln */
    function changeView(ansicht)
    {
        window.location.href="alarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=datum%>&filter=nein&suche=nein&monat=<%=aktMonat%>&jahr=<%=aktJahr%>&ansicht="+ansicht+"";
    }

    /* Druckbare Seite öffnen */
    function loadPrint()
    {
        <% If meldeArt2 = "1" then 'Historische Alarme %>
            if(document.forms["Druckansicht"])
            {
                document.forms["Druckansicht"].elements["sql"].value="<%=SQLStmt%>";
                document.forms["Druckansicht"].elements["meldeart"].value="<%=meldeArt2%>";
                document.forms["Druckansicht"].elements["kategorie"].value="<%=meldeKat2%>";
                document.forms["Druckansicht"].elements["datum"].value="<%=aktDatum%>";
                document.forms["Druckansicht"].elements["filter"].value="<%=filter%>";
                document.forms["Druckansicht"].elements["filterText"].value="<%=filterText%>";
                document.forms["Druckansicht"].elements["spalte"].value="<%=spalte%>";
                document.forms["Druckansicht"].elements["suche"].value="<%=suche%>";
                document.forms["Druckansicht"].elements["suchText2"].value="<%=suchText%>";
                document.forms["Druckansicht"].elements["monat"].value="<%=aktMonat%>";
                document.forms["Druckansicht"].elements["jahr"].value="<%=aktJahr%>";
                document.forms["Druckansicht"].elements["zeit"].value="<%=zeitString%>";
                document.forms["Druckansicht"].elements["von"].value="<%=letzteDatensaetze-recordAnzCount%>";
                document.forms["Druckansicht"].elements["bis"].value="<%=letzteDatensaetze%>";
                document.forms["Druckansicht"].elements["intMaxRecord"].value="<%=intMaxRecord%>";
                document.forms["Druckansicht"].elements["bTopDown"].value="<%=bTopDown%>";

                document.forms["Druckansicht"].submit();
            }
        <% Else %>
            window.open("printAlarme.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=aktDatum%>&filter=<%=filter%>&filterText=<%=ChangeUmlInhtml2(filterText)%>&spalte=<%=spalte%>&suche=<%=suche%>&suchText2=<%=ChangeUmlInhtml2(suchText)%>&monat=<%=aktMonat%>&jahr=<%=aktJahr%>&zeit=<%=zeitString%>&von=<%=letzteDatensaetze-recordAnzCount%>&bis=<%=letzteDatensaetze%>&intMaxRecord=<%=intMaxRecord%>&bTopDown=<%=bTopDown%>&sql=<%=ChangeUmlInhtml2(SQLStmt)%>","_blank","menubar=no, status=no, toolbar=no, scrollbars=yes, width=800, height=600");
        <% End If %>
    }

    /* Dateiexport */
    function exportCsv()
    {
        <% If meldeArt2 = "1" then 'Historische Alarme %>
            document.forms["ExportForm"].elements["sql"].value="<%=SQLStmt%>";
            document.forms["ExportForm"].elements["meldeart"].value="<%=meldeArt2%>";
            document.forms["ExportForm"].elements["kategorie"].value="<%=meldeKat2%>";
            document.forms["ExportForm"].elements["datum"].value="<%=aktDatum%>";
            document.forms["ExportForm"].elements["filter"].value="<%=filter%>";
            document.forms["ExportForm"].elements["filterText"].value="<%=filterText%>";
            document.forms["ExportForm"].elements["spalte"].value="<%=spalte%>";
            document.forms["ExportForm"].elements["suche"].value="<%=suche%>";
            document.forms["ExportForm"].elements["suchText2"].value="<%=suchText%>";
            document.forms["ExportForm"].elements["monat"].value="<%=aktMonat%>";
            document.forms["ExportForm"].elements["jahr"].value="<%=aktJahr%>";
            document.forms["ExportForm"].elements["zeit"].value="<%=zeitString%>";
            document.forms["ExportForm"].elements["von"].value="<%=letzteDatensaetze-recordAnzCount%>";
            document.forms["ExportForm"].elements["bis"].value="<%=letzteDatensaetze%>";
            document.forms["ExportForm"].elements["intMaxRecord"].value="<%=intMaxRecord%>";
            document.forms["ExportForm"].elements["bTopDown"].value="<%=bTopDown%>";

            document.forms["ExportForm"].submit();
        <% Else %>
            window.open("alarmExport.asp?meldeart=<%=meldeArt2%>&kategorie=<%=meldeKat2%>&datum=<%=aktDatum%>&filter=<%=filter%>&filterText=<%=ChangeUmlInhtml2(filterText)%>&spalte=<%=spalte%>&suche=<%=suche%>&suchText2=<%=ChangeUmlInhtml2(suchText)%>&monat=<%=aktMonat%>&jahr=<%=aktJahr%>&zeit=<%=zeitString%>&von=<%=letzteDatensaetze-recordAnzCount%>&bis=<%=letzteDatensaetze%>&intMaxRecord=<%=intMaxRecord%>&bTopDown=<%=bTopDown%>&sql=<%=ChangeUmlInhtml2(SQLStmt)%>","_blank","menubar=yes, status=no, toolbar=no, scrollbars=yes, resizable=yes width=680, height=600");
        <% End If%>
    }

    /* Zeilen beim Anklicken farbig markieren --> Wenn Benutzer auf Info-Button in der Zeile
    klickt, Zeile farbig kennzeichnen um Übersicht über angesehene Meldungen zu erhöhen */
    var oldColor = "-";
    var oldIndex = "-";
    function setRowColor(oldC,index)
    {
        if (oldColor != "-" && oldIndex != "-")
        {
            if(oldIndex != index)
                window.document.getElementById("Zeile" + oldIndex).className = oldColor;
        }
        if (oldIndex != index)
        {
            window.document.getElementById("Zeile" + index).className = "rowClickColor";
            oldColor=oldC;
            oldIndex=index;
        }
    }

    /* ------------------------------------------------------------------------------------------------- */
    /* Aufruf des Properties-Fenster des jeweiligen Alarms */
    function details(nam, ziel, wid, hei)
    {
        xx=ziel;
        Details1 = window.open (xx,nam,"width="+wid+",height="+hei+",top=100, left=100, resizable=yes,locationbar=no,menubar=no,scrollbars=no,status=no,toolbar=no");
    }
    /* ------------------------------------------------------------------------------------------------- */

    /*--------------------------------------------------------------------------------------------*/
    /*-- Formatierungsfunktionen                                                                --*/
    /*--------------------------------------------------------------------------------------------*/
    function ChangeString(BearbeitungsString, ZeichenAlt, ZeichenNeu)
    {
        var temp1, temp2;
        var posA, ZeichenAltLen;

        ZeichenAltLen = ZeichenAlt.length;

        for(var i=0; i< BearbeitungsString.length; i++)
        {
            posA = BearbeitungsString.indexOf(ZeichenAlt);
            if(posA >= 0)
            {
                temp1 = BearbeitungsString.substring(0,posA);
                temp2 = BearbeitungsString.substring(posA + ZeichenAltLen, BearbeitungsString.length);
                BearbeitungsString = temp1 + ZeichenNeu + temp2;
            }
            else
            {
                break;
            }
        }
        return BearbeitungsString;
    }


    /* Url - Formatierungsfunktion für Parameterübergabe über Url´s */
    function ChangeUmlInhtml2(UeString)
    {
        var TempPath;
        TempPath = UeString;

        //TempPath = ChangeString(TempPath, "\\", "[~A1~]");
        TempPath = ChangeString(TempPath, "/", "[~A2~]");
        TempPath = ChangeString(TempPath, ":", "[~A3~]");
        TempPath = ChangeString(TempPath, "*", "[~A4~]");
        TempPath = ChangeString(TempPath, "?", "[~A5~]");
        TempPath = ChangeString(TempPath, "\"", "[~A6~]");
        TempPath = ChangeString(TempPath, ">", "[~A7~]");
        TempPath = ChangeString(TempPath, "<", "[~A8~]");
        TempPath = ChangeString(TempPath, "|", "[~A9~]");

        TempPath = ChangeString(TempPath, " ", "[~B1~]");
        TempPath = ChangeString(TempPath, "Ä", "[~B2~]");
        TempPath = ChangeString(TempPath, "ä", "[~B3~]");
        TempPath = ChangeString(TempPath, "Ö", "[~B4~]");
        TempPath = ChangeString(TempPath, "ö", "[~B5~]");
        TempPath = ChangeString(TempPath, "Ü", "[~B6~]");
        TempPath = ChangeString(TempPath, "ü", "[~B7~]");
        TempPath = ChangeString(TempPath, "ß", "[~B8~]");

        TempPath = ChangeString(TempPath, "#", "[~B9~]");
        TempPath = ChangeString(TempPath, "'", "[~C1~]");
        TempPath = ChangeString(TempPath, "%", "[~C2~]");
        TempPath = ChangeString(TempPath, "&", "[~C3~]");
        TempPath = ChangeString(TempPath, "+", "[~C4~]");
        TempPath = ChangeString(TempPath, ".", "[~C5~]");

        return TempPath;
    }

    /* Open data point detail window */
    function openDetails(userAdr)
    {
        // Fenster zentrieren
        var fensterBreite2;
        fensterBreite2 = 1024;
        posX = (screen.width - fensterBreite2) / 2;
        fensterHoehe2 = 515;
        posY = (screen.height - fensterHoehe2 - 100) / 2;
        fenster1 = window.open(
                "/webvisionnt/zsystem/maindata/detail.asp?useradr=" + userAdr,
                "_blank",
                "top=" + posY + ", left=" + posX + ", width=" + fensterBreite2 + ",height=" + fensterHoehe2 +
                ",resizable=yes, toolbar=no ,menubar=no, scrollbars=no");
    }

    /* Meldemanager - Hauptfenster öffnen */
    function openMeldeManager(userAdr)
    {
        F1 = window.open("Meldemanager/mmFrameHaupt.asp?Wert=-&sql=" + userAdr + "&prioline=1&prioFilter=<%=prio %>", 
                            "_blank", "width=950, height=750, menubar=no, resizable=no, scrollbars=no, toolbar=no," +
                            " top=10, left=100");
    }

    function RefreshLogo()
	{
		if(window.parent.Oben)
		{
			if(window.parent.Oben.document.getElementById("ModulLogo"))
			{
				//alert("<%=Session("MMAlarmStatus")%>");
				window.parent.Oben.SetLogoMode(<%=Session("MMAlarmStatus")%>);
			}
		}
	}

</script>
</body>
</html>