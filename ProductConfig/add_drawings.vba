Public Sub add_drawings()
    'Die Prozedur importiert einen Bestand an Einbauzeichnungen und Gegenflanschen aus dem Verzeichnis "_Import" in die Tabelle "gearbox_files".
    'mit Hilfe der erforderlichen Kontrolldateien "245_Einbauzeichungen.csv" und "220_Gegenflansche.csv"
    'Dabei werden kennzeichnende Merkmale zu einer Beschreibung zusammengefasst wie auch die Dateiimporte gesteuert.
    'ACHTUNG! Bearbeitung der Kontrolldateien erforderlich:
    '- in der Kopfzeile der CSV-Dateien "[mm]" durch "(mm)" ersetzen
    '- Dateien in lokale Tabellen importieren
    '- Baugrößen mit Tabelle 'gearboxes' abgleichen und ggf. ändern (z.B. 663 -> 665, VLJ 430/1 -> VLJ 430) oder löschen (z.B. WAF 144)
    '- ggf. RHS Datensätze anpassen
    '- ggf. Fundamentpläne kopieren und angepassen
    'Nach dem Import: Kopien der WAF-Datensätze für LAF anlegen

    'Alle vorhandenen Datensätze in "gearbox_files" und "gearbox_file_mappings" werden ohne Backup gelöscht.

    On Error GoTo ErrHandler
 Deklarationen:
    Const sProcName As String = "add_drawings"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String, sUser As String
    Dim sPfad(2), sFNam(2) As String, sExt As String, sPref As String, sDraw As String, sTmp As String, sPNam(2)
    Dim sCtlTab As String, sTab(4) As String, sMime(2) As String, sConn As String, sDsc As String
    Dim sBAR As String, sBGR As String, sBAF As String, sPTO As String, sMNT As String, sCtl As String, sSAE As String
    Dim sPTR As String, sWBR As String
    Dim i1 As Long, imax As Long, lFID As Long
    Dim iMot As Integer, i As Integer
    Dim vAnt As Variant
    
    Dim rsCtl As DAO.Recordset
    Dim rs(2) As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim myStream(2) As ADODB.Stream
    Dim fso As FileSystemObject
    Dim oOrdner As Object
    
    sMime(1) = "application/pdf"
    sMime(2) = "application/zip"
    sTab(1) = "gearbox_files"
    sTab(2) = "gearbox_file_mappings"
    
 Bestaetigung:
    sMsg = "This procedure will import a complete new set of installation drawings for " _
         & "REINTJES Product Configurator." & vbLf _
         & "Requirements: CSV control tables and corresponding drawings (PDF, ZIP) in directory '_Import' " & vbLf & vbLf _
         & "Unfortunatly DDL commands (used to reset auto_increment) for non-Access tables are not supported. " & vbLf & vbLf _
         & "YES: INSERT new drawings without deleting existing records. " & vbLf & vbLf _
         & "NO: DELETE existing records and QUIT the procedure without inserting new files. " _
         & "ATTENTION! Existing records in tables '" & sTab(1) & "' and '" & sTab(2) & "' " _
         & "will be deleted without backup!"
    vAnt = MsgBox(sMsg, vbExclamation + vbYesNoCancel, sProcName)
    
    Select Case vAnt
        Case vbCancel
            Exit Sub
        Case vbNo
            'Zieltabelle wird geleert (über Foreign_Keys verbundene Tabellen ebenfalls)
            sSQL = "DELETE " & sTab(1) & ".* FROM " & sTab(1) & " ;"
            DoCmd.SetWarnings False
            DoCmd.RunSQL sSQL
            sMsg = "Records in 'reintjes." & sTab(1) & " deleted." & vbLf _
                 & "To reset AUTO_INCREMENT for the file_ID manually use " & vbLf _
                 & "'ALTER TABLE reintjes.gearbox_files AUTO_INCREMENT=1;' statement in external application."
            Debug.Print sMsg
            MsgBox sMsg, vbInformation, sProcName
            Exit Sub
        Case vbYes
            sPfad(1) = GetFolderDialog("M:\TN\Normung\Anwendungsdaten\Access2010\ProductGuide\_Import\")
            Set fso = CreateObject("Scripting.FileSystemObject")
    End Select
    
 MySQLconnect:
    'Neue ADODB-Connection zu MySQL-Datenbank initialisieren
    'functional: DRIVER={MySQL ODBC 5.1 Driver}
    'to be testet: DRIVER={MySQL ODBC 5.3 Unicode Driver}
    sConn = "DRIVER={MySQL ODBC 5.3 Unicode Driver}; " & _
            "SERVER=rpksrv.reintjes.loc; " & _
            "DATABASE=reintjes; " & _
            "User Id=root; " & _
            "Password=mei_cei1ai7phahD; " & _
            "OPTION=16427"
    Set oConn = New ADODB.Connection
    oConn.Open sConn
    oConn.CursorLocation = adUseClient
    
    Set rs(1) = New ADODB.Recordset
    sSQL = "SELECT * FROM " & sTab(1) & " WHERE 1=0"
    rs(1).Open sSQL, oConn, adOpenStatic, adLockOptimistic
    
    Set rs(2) = New ADODB.Recordset
    sSQL = "SELECT * FROM " & sTab(2) & " WHERE 1=0"
    rs(2).Open sSQL, oConn, adOpenStatic, adLockOptimistic
    
    ' Beide Zieltabellen werden parallel gefüllt, da die korrespondierende File_ID auch in der
    ' mappings-Tabelle hinterlegt werden muss.

 Einbauzeichnungen:
    sCtlTab = "245_Einbauzeichungen"
    Set rsCtl = CurrentDb.OpenRecordset(sCtlTab, dbOpenSnapshot)
    i1 = 0
    
    With rsCtl
        'Show status bar
        .MoveLast
        imax = .RecordCount
        .MoveFirst
        SysCmd acSysCmdInitMeter, "Please wait! Importing drawings ...", imax
        
        Do While Not .EOF
            i1 = i1 + 1
            SysCmd acSysCmdUpdateMeter, i1
            
            'Dateinamen
            sPNam(1) = sPfad(1) & Nz(Dir(sPfad(1) & CStr(!ID) & "*.pdf"))
            sFNam(1) = fso.GetFileName(sPNam(1))
            
            sPNam(2) = sPfad(1) & Nz(Dir(sPfad(1) & CStr(!ID) & "*.zip"))
            sFNam(2) = fso.GetFileName(sPNam(2))
            
            'Bezeichnung
            sPref = Nz(![Vor-Nr])
            sDraw = sPref & CStr(!ID)
            
            Select Case sPref
                Case "0-104-"
                    sDsc = "Installation"
                Case "0-107-"
                    Select Case !Bezeichnung
                        Case "Einbauskizze"
                            sDsc = "Dimension sheet"
                        Case "Fundamentplan"
                            sDsc = "Foundation plan"
                        Case Else
                    End Select
                Case Else
            End Select
            
            'Baureihe / Series
            sBAR = Nz(!Getr_Typ)
            
            'Baugröße / Size
            sTmp = CStr(Nz(!GetrGr_von))
            sBGR = IIf(InStr(Nz(!Getr_Zus), "/1") > 0, sTmp & "/1", sTmp)
            
            'Bauform / Design und Achslage
            sTmp = Nz(!Achslage)
            Select Case sTmp
                Case "vertikal": sBAF = "V"
                Case "horizontal links": sBAF = "H"
                Case "horizontal rechts": sBAF = "H"
                Case "diagonal links": sBAF = "D"
                Case "diagonal rechts": sBAF = "D"
                Case Else: sBAF = "V"
            End Select
            
            'Hybrid-Systeme
            iMot = 0
            sTmp = Nz(!Getr_Zus)
            If InStr(sTmp, "RHS") > 0 Then
                sBAF = "HS"
                If InStr(Nz(!Z_Bemerkung), "60") > 0 Then iMot = 60
                If InStr(Nz(!Z_Bemerkung), "100") > 0 Then iMot = 100
                If InStr(Nz(!Z_Bemerkung), "200") > 0 Then iMot = 200
                If InStr(Nz(!Z_Bemerkung), "315") > 0 Then iMot = 315
                If InStr(Nz(!Z_Bemerkung), "400") > 0 Then iMot = 400
                If InStr(Nz(!Z_Bemerkung), "500") > 0 Then iMot = 500
                If InStr(Nz(!Z_Bemerkung), "630") > 0 Then iMot = 630
            End If
            
            sTmp = Nz(![Nebentrieb 1])
            sPTO = IIf(Left(sTmp, 1) = "K", Left(sTmp, 3), "")
            
            sTmp = Nz(!Aufstellung)
            Select Case sTmp
                Case "starr": sMNT = "rigid"
                Case "ohne Fußwinkel": sMNT = "no brackets"
                Case "starr m. Fusswinkel": sMNT = "rigid, brackets"
                Case Else: sMNT = sTmp
            End Select
            
            sTmp = Nz(!Steuerung)
            Select Case sTmp
                Case "elektrisch": sCtl = "electrical"
                Case "mechanisch": sCtl = "mechanical"
                Case Else: sCtl = sTmp
            End Select
            
            sTmp = Trim(Left(Nz(!Zwischengehaeuse), 5))
            Select Case sTmp
                Case "SAE0", "SAE00", "SAE1": sSAE = sTmp
                Case "ohne": sSAE = ""
                Case Else: sSAE = sTmp
            End Select
            
            sWBR = IIf(Nz(!Wellenbremse) = "Ja", "ShBr", "")
            sPTR = IIf(Nz(!Pumpentrieb) = "Einfach", "PDrv", "")
            sDsc = sDsc & "; " & sBAR & " " & sBGR & " " & sBAF _
                 & IIf(iMot > 0, CStr(iMot / 10), "") _
                 & IIf(Len(sPTO) > 0, " " & sPTO, "") _
                 & IIf(Len(sSAE) > 0, " " & sSAE, "") & ";" _
                 & IIf(Len(sMNT) > 0, " Mount: " & sMNT, "") & ";" _
                 & IIf(Len(sCtl) > 0, " Ctl: " & sCtl, "") & " " _
                 & " " & sWBR & " " & sPTR
            sDsc = Trim(sDsc)
            
            'Anfügen neuer Datensätze für Einbauzeichnungen
            'i: 1= PDF,  2= ZIP
            
            For i = 1 To 2
                If Not fso.FileExists(sPNam(i)) Then
                    sMsg = "File not found: " & sFNam(i)
                    GoTo NextStep104
                End If
                Set myStream(i) = New ADODB.Stream
                myStream(i).Type = adTypeBinary
                myStream(i).Open
                myStream(i).LoadFromFile sPNam(i)
                
                With rs(1)  'files
                    .AddNew
                    ![Name] = sFNam(i)
                    ![mime] = sMime(i)
                    ![Size] = myStream(i).Size
                    ![Data] = myStream(i).Read
                    ![Description] = sDsc
                    .Update
                    lFID = ![ID]
                End With
                
                With rs(2)  'mapping files
                    .AddNew
                    ![gearbox_series] = sBAR
                    ![gearbox_size] = sBGR
                    ![gearbox_design] = sBAF
                    ![gearbox_e_motor] = iMot
                    ![file_id] = lFID
                    .Update
                End With
                sMsg = "File successfully inserted: " & sFNam(i)
            
                myStream(i).Close
                Set myStream(i) = Nothing
 NextStep104:
                Debug.Print CStr(i1) & "-" & CStr(lFID), sMsg, sDsc
            Next i
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsCtl = Nothing
    
    SysCmd acSysCmdRemoveMeter
    DoCmd.OpenQuery "A_Zeichnungen_LAF_add"
    sMsg = CStr(rs(1).RecordCount) & " records for installation drawings written into 'gearbox_files' " _
         & "and into 'gearbox_file_mappings'."
    Debug.Print sMsg
    MsgBox sMsg, vbInformation, sProcName
    
 Flansche:
    sCtlTab = "220_Gegenflansche"
    Set rsCtl = CurrentDb.OpenRecordset(sCtlTab, dbOpenSnapshot)
    i1 = 0
    
    With rsCtl
        'Show status bar
        .MoveLast
        imax = .RecordCount
        .MoveFirst
        SysCmd acSysCmdInitMeter, "Please wait! Importing drawings ...", imax
        
        Do While Not .EOF
            i1 = i1 + 1
            SysCmd acSysCmdUpdateMeter, i1
            
            'Baureihe / Series
            sBAR = Nz(!Getr_Typ)
            
            'Baugröße / Size
            sTmp = CStr(Nz(!GetrGr_von))
            sBGR = IIf(InStr(Nz(!Getr_Zus), "/1") > 0, sTmp & "/1", sTmp)
            
            'Bauform / Design
            sBAF = "V"
            
            'E-Motor
            iMot = 0
            
            'Dateinamen
            sPNam(1) = sPfad(1) & Nz(Dir(sPfad(1) & CStr(!ID) & "*.pdf"))
            sFNam(1) = fso.GetFileName(sPNam(1))
            
            sPNam(2) = sPfad(1) & Nz(Dir(sPfad(1) & CStr(!ID) & "*.zip"))
            sFNam(2) = fso.GetFileName(sPNam(2))
            
            sDsc = "Counter flange" & "; " & sBAR & " " & sBGR
            
            'Anfügen neuer Datensätze für Gegenflansche
            'i: 1= PDF,  2= ZIP
            
            For i = 1 To 2
                If Not fso.FileExists(sPNam(i)) Then
                    sMsg = "File not found: " & sFNam(i)
                    GoTo NextStep202
                End If
                Set myStream(i) = New ADODB.Stream
                myStream(i).Type = adTypeBinary
                myStream(i).Open
                myStream(i).LoadFromFile sPNam(i)
                
                With rs(1)  'files
                    .AddNew
                    ![Name] = sFNam(i)
                    ![mime] = sMime(i)
                    ![Size] = myStream(i).Size
                    ![Data] = myStream(i).Read
                    ![Description] = sDsc
                    .Update
                    lFID = ![ID]
                End With
                
                With rs(2)  'mapping files
                    .AddNew
                    ![gearbox_series] = sBAR
                    ![gearbox_size] = sBGR
                    ![gearbox_design] = sBAF
                    ![gearbox_e_motor] = iMot
                    ![file_id] = lFID
                    .Update
                End With
                sMsg = "File successfully inserted: " & sFNam(i)

                myStream(i).Close
                Set myStream(i) = Nothing
 NextStep202:
                Debug.Print CStr(i1) & "-" & CStr(lFID), sMsg, sDsc
            Next i
            
            .MoveNext
        Loop
        .Close
    End With
    
    SysCmd acSysCmdRemoveMeter
    sMsg = CStr(rs(1).RecordCount) & " records for counter flanges (PDF+ZIP) written into tables 'gearbox_files' " _
         & "and 'gearbox_file_mappings'."
    Debug.Print sMsg
    MsgBox sMsg, vbInformation, sProcName
    
    rs(1).Close
    rs(2).Close
    Set rs(1) = Nothing
    Set rs(2) = Nothing
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print "Current file: " & sFNam(i) & " - " & sDsc
        Debug.Print "Current mapping: " & sBAR & "_" & sBGR & "_" & sBAF & "_" & CStr(iMot) & "_" & CStr(lFID)
        Debug.Print sMsg
    End With
    Resume ExitHere
End Sub