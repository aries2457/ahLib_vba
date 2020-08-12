Option Compare Database
Option Explicit

Const sModName As String = "mod_Insert"

Public Sub add_gearbox_series(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_series"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_series"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Baureihen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearbox_series " & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    DoCmd.SetWarnings False
    
    'Bestehende Einträge für Baureihen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Baureihen neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_series ( `name`, `properties`, `note` ) " _
         & "SELECT T_gearbox_series.`name`, T_gearbox_series.`properties`, T_gearbox_series.`note` " _
         & "FROM T_gearbox_series ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_option_localization(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_option_localization"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    
    sTab0 = "T_Preiskomponenten"
    sTab1 = "gearbox_option_localization"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Sprachen für Getriebeoptionen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): " & sTab0 & vbLf _
         & "Voraussetzungen: gearbox_options " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Sprachen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Sprache 'en-US' neu einfügen
    sSQL = "INSERT INTO gearbox_option_localization ( option_name, `language`, localization ) " _
         & "SELECT " & sTab0 & ".KPID, 'en-US' AS lang, " & sTab0 & ".Bez_en " _
         & "FROM " & sTab0 & " INNER JOIN gearbox_options ON " & sTab0 & ".KPID = gearbox_options.`name` " _
         & "ORDER BY " & sTab0 & ".KPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
    'Einträge für Sprache 'de-DE' neu einfügen
    sSQL = "INSERT INTO gearbox_option_localization ( option_name, `language`, localization ) " _
         & "SELECT " & sTab0 & ".KPID, 'de-DE' AS lang, " & sTab0 & ".Bez_de " _
         & "FROM " & sTab0 & " INNER JOIN gearbox_options ON " & sTab0 & ".KPID = gearbox_options.`name` " _
         & "ORDER BY " & sTab0 & ".KPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_options(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_options"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    
    sTab0 = "T_Preiskomponenten"
    sTab1 = "gearbox_options"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Optionen für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): " & sTab0 & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    
    'Einträge für societies neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_options ( `name` ) " _
         & "SELECT " & sTab0 & ".KPID " _
         & "FROM " & sTab0 & " " _
         & "WHERE " & sTab0 & ".KPID NOT LIKE '0*' " _
         & "ORDER BY " & sTab0 & ".KPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_option_prices(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_option_prices"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    
    sTab0 = "T_Preise"
    sTab1 = "gearbox_option_prices"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Preise für Optionen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): " & sTab0 & vbLf _
         & "Voraussetzungen: gearbox_options " & vbLf _
         & "Nachbearbeitung: --." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für OptionsPreise neu einfügen (aktuell auf Standard-Typen begrenzt)
    sSQL = "INSERT INTO " & sTab1 & " ( series, `size`, `option`, price ) " _
         & "SELECT " & sTab0 & ".BAR, " & sTab0 & ".BGR, " & sTab0 & ".KPID, " & sTab0 & ".Preis1 " _
         & "FROM (T_StdTypen INNER JOIN " & sTab0 & " ON (T_StdTypen.BAR = " & sTab0 & ".BAR) " _
         & "AND (T_StdTypen.BGR = " & sTab0 & ".BGR) AND (T_StdTypen.BAF = " & sTab0 & ".BAF)) " _
         & "INNER JOIN gearbox_options ON " & sTab0 & ".KPID = gearbox_options.name " _
         & "WHERE (((" & sTab0 & ".Preis1) Is Not Null) And ((" & sTab0 & ".MOT) Is Null)) " _
         & "ORDER BY " & sTab0 & ".BAR, " & sTab0 & ".BGR, " & sTab0 & ".BAF, " & sTab0 & ".KPID ; "
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_attribute_descriptions(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_attribute_descriptions"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_attribute_descriptions"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Bemaßungsattribute für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Attribute" & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für societies neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_attribute_descriptions ( attribute_key, unit, comment ) " _
         & "SELECT T_Attribute.ATNAM, T_Attribute.ATDIM, T_Attribute.NOTE " _
         & "FROM T_Attribute " _
         & "ORDER BY T_Attribute.ATNAM ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_classification_societies(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_classification_societies"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_classification_societies"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Classification Societies ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearbox_classification_societies " & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    
    'Einträge für societies neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_classification_societies ( `name` ) " _
         & "SELECT T_gearbox_classification_societies.`name` " _
         & "FROM T_gearbox_classification_societies;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_engine_manufacturers(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_engine_manufacturers"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "engine_manufacturers"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Motorenhersteller ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_MotHersteller " & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Motorenhersteller neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO engine_manufacturers ( `name`, description ) " _
         & "SELECT T_MotHersteller.Maker, T_MotHersteller.LName " _
         & "FROM T_MotHersteller " _
         & "WHERE (((T_MotHersteller.Maker) Not Like '_*')) ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_gearboxes(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearboxes"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearboxes"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Getriebetypen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearboxes " & vbLf _
         & "Voraussetzungen: gearbox_series gefüllt " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Getriebetypen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    
   'Einträge für Getriebetypen neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearboxes ( series, `size`, design, e_motor, comment ) " _
         & "SELECT T_gearboxes.series, T_gearboxes.`size`, T_gearboxes.design, T_gearboxes.e_motor, T_gearboxes.comment " _
         & "FROM T_gearboxes ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_classification_classes(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_classification_classes"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_classification_classes"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für DutyClasses ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearbox_classification_classes " & vbLf _
         & "Voraussetzungen: gearbox_classification_societies gefüllt " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Getriebetypen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    
   'Einträge für Getriebetypen neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_classification_classes ( `name`, classification_society ) " _
         & "SELECT T_gearbox_classification_classes.`name`, T_gearbox_classification_classes.classification_society " _
         & "FROM T_gearbox_classification_classes ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_power_moments(Optional bClean As Boolean = True)
    ' Purpose:  Add gearbox P2S ratios according to class
    ' Author:   Andreas Herrel
    ' Date:     2018-08-28  updated: 2018-12-13
    ' Inputs:   ProductGuideS-BE(A_LM_Katalog,T_Untersetzungen), T_LM_Fieldmappings
    ' Output:   gearbox_power_moments
    ' Requires: gearbox_classification_classes, gearbox_prices

    On Error GoTo ErrHandler
    Const sProcName As String = "add_power_moments"
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sClass As String, sFeld1 As String
    
    sTab1 = "gearbox_power_moments"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Leistungsmomente für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideX-BE (T_LM_Katalog_X, T_Untersetzungen) " & vbLf _
         & "Voraussetzungen: gearbox_classification_classes + ...societies gefüllt " & vbLf _
         & "Nachbearbeitung: gearbox_power_moment_mappings neu füllen" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge löschen
    DoCmd.SetWarnings False
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Leistungmomente neu einfügen (Quelle ProductGuideS-DB)
    'Andere Datenstruktur erfordert mehrfaches Anfügen
    Set rs1 = CurrentDb.OpenRecordset("T_LM_Fieldmappings", dbOpenSnapshot)
    
    With rs1
        Do While Not .EOF
        'DutyClasses
            sClass = !Class
            sFeld1 = !FldName
            sSQL = "INSERT INTO " & sTab1 & " ( gearbox_series, gearbox_size, Ratio, class, power, power_moment, special ) " _
                 & "SELECT A_LM_Katalog.Baureihe, A_LM_Katalog.GetriebeGr, A_LM_Katalog.Ratio, '" & sClass & "' AS class, " _
                 & "Nz([T_Untersetzungen].[Leistung_kW],0) AS power, A_LM_Katalog." & sFeld1 & ", IIf(T_Untersetzungen.Knz In ('#','i'),1,0) AS special " _
                 & "FROM A_LM_Katalog LEFT JOIN T_Untersetzungen ON (A_LM_Katalog.Ratio = T_Untersetzungen.Ratio) AND " _
                 & "(A_LM_Katalog.GetriebeGr = T_Untersetzungen.GetriebeGr) AND (A_LM_Katalog.Baureihe = T_Untersetzungen.Baureihe) " _
                 & "WHERE (((A_LM_Katalog.Baureihe) In ('LAF','VLJ','WAF','WLS','WVS')) AND ((A_LM_Katalog." & sFeld1 & ") Is Not Null) " _
                 & "AND ((Val([A_LM_Katalog]![GetriebeGr])) Between 100 And 2000)) " _
                 & "OR (((A_LM_Katalog.Baureihe)='VLJ') AND ((A_LM_Katalog." & sFeld1 & ") Is Not Null) " _
                 & "AND ((Val([A_LM_Katalog]![GetriebeGr])) Between 100 And 1000)) " _
                 & "ORDER BY A_LM_Katalog.Baureihe, A_LM_Katalog.GetriebeGr;"
            Debug.Print sClass, sSQL
            DoCmd.RunSQL sSQL
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs1 = Nothing
    DoCmd.SetWarnings True
    sSQL = "SELECT " & sTab1 & ".* FROM " & sTab1 & " ;"
    Set rs1 = CurrentDb.OpenRecordset(sSQL, dbOpenSnapshot)
    With rs1
        .MoveLast
        sMsg = "Procedure completed normally. " & vbCrLf & CStr(Nz(.RecordCount)) & " records inserted into " & sTab1
        .Close
    End With
    Set rs1 = Nothing
    MsgBox sMsg, vbInformation, sProcName

 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: Modul" & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
End Sub

Public Sub add_engines(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_engines"
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sClass As String, sFeld1 As String
    
    sTab1 = "engines"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Dieselmotoren ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Motoren " & vbLf _
         & "Voraussetzungen: engine_manufacturers gefüllt " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge in Zieltabelle löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Leistungmomente für Motoren neu einfügen
    sSQL = "INSERT INTO engines ( manufacturer, Type, rating_category, rating, Speed, Torque, power, Series, `application`, `Note` ) " _
         & "SELECT T_Motoren.Maker, T_Motoren.Type, T_Motoren.RatingCat, T_Motoren.Rating, T_Motoren.Speed, T_Motoren.Torque, T_Motoren.Rating, T_Motoren.Series, T_Motoren.Appl, T_Motoren.Note " _
         & "FROM T_Motoren " _
         & "WHERE (((T_Motoren.Maker) Not Like '_*')) " _
         & "ORDER BY T_Motoren.Maker, T_Motoren.Type ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL

 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: Modul" & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_attribute_mappings(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_attribute_mappings"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_attribute_mappings"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt Mappings für neue Bemaßungen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Abmessungen_x" & vbLf _
         & "Voraussetzungen: gearbox_attribute_descriptions & ...mappings gefüllt " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Abmessungen von Standard-Getriebeypen neu einfügen (Quelle ProductGuideS)
    sSQL = "INSERT INTO " & sTab1 & " ( gearbox_series, gearbox_size, gearbox_design, attribute_key, gearbox_e_motor ) " _
         & "SELECT T_StdTypen.BAR, T_StdTypen.BGR, T_StdTypen.BAF, T_Abmessungen_x.ATNAM, 0 AS MOT " _
         & "FROM T_Abmessungen_x INNER JOIN T_StdTypen ON T_Abmessungen_x.TYPID = T_StdTypen.TYPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

Public Sub add_attributes(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_attributes"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_attributes"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Bemaßungen für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Abmessungen_x" & vbLf _
         & "Voraussetzungen: gearbox_attribute_descriptions & ...mappings gefüllt " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Attribute löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Abmessungen von Standard-Getriebeypen neu einfügen (Quelle ProductGuideS)
    sSQL = "INSERT INTO " & sTab1 & " ( gearbox_series, gearbox_size, gearbox_design, attribute_key, attribute_value ) " _
         & "SELECT T_StdTypen.BAR, T_StdTypen.BGR, T_StdTypen.BAF, T_Abmessungen_x.ATNAM, T_Abmessungen_x.ATVALN " _
         & "FROM T_Abmessungen_x INNER JOIN T_StdTypen ON T_Abmessungen_x.TYPID = T_StdTypen.TYPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

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

Public Sub add_gearbox_descriptions(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_descriptions"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sTab1 As String
    
    sTab1 = "gearbox_descriptions"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Bemaßungen für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Abmessungen_x" & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Neue Kurzbeschreibungen für Getriebe hinzufügen
    sSQL = "INSERT INTO " & sTab1 & " ( id, `language`, description ) " _
         & "SELECT T_gearbox_descriptions.id, T_gearbox_descriptions.lang, T_gearbox_descriptions.note " _
         & "FROM T_gearbox_descriptions " _
         & "ORDER BY T_gearbox_descriptions.id, T_gearbox_descriptions.lang ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
        
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere
End Sub

Public Sub add_gearbox_description_mappings(Optional bClean As Boolean = True)
    ' Purpose:  Add description mappings for gearbox series
    ' Author:   Andreas Herrel
    ' Date:     2018-10-15; updated: 2018-10-17
    ' Inputs:   gearboxes; gearbox_descriptions; T_gearbox_descriptions
    ' Output:   gearbox_description_mappings
    ' Requires: MySQL database online

    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_description_mappings"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sTab() As String
    
    sTab(1) = "gearbox_description_mappings"
    sTab(2) = "T_gearbox_descriptions"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Beschreibungen für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab(1) & vbLf _
         & "Quelle(n): gearboxes; gearbox_descriptions; T_gearbox_descriptions" & vbLf _
         & "Voraussetzungen: Quellen vorhanden " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab(1) & ".* FROM " & sTab(1) & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Neue Kurzbeschreibungen für Getriebetypen hinzufügen
    sSQL = "INSERT INTO " & sTab(1) & " ( gearbox_series, gearbox_size, gearbox_design, " _
         & "gearbox_e_motor, description_id, description_language ) " _
         & "SELECT gearboxes.series, gearboxes.size, gearboxes.design, gearboxes.e_motor, " _
         & "gearbox_descriptions.id, gearbox_descriptions.language " _
         & "FROM (gearbox_descriptions INNER JOIN " & sTab(2) & " ON " _
         & "(gearbox_descriptions.language = " & sTab(2) & ".lang) AND " _
         & "(gearbox_descriptions.id = " & sTab(2) & ".id)) INNER JOIN gearboxes " _
         & "ON " & sTab(2) & ".series = gearboxes.series " _
         & "ORDER BY gearboxes.series, gearboxes.size, gearboxes.design ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
        
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere
End Sub

Public Sub add_auxiliary_drive_prices(Optional bClean = True)
    ' Purpose: Add basic prices for auxiliary drives
    ' Author: Andreas Herrel
    ' Date: 2018-10-16;  last update: 2018-11-23
    ' Inputs: gearbox_prices; T_Preise
    ' Output: auxiliary_drive_prices
    ' Requirements: MySQL database online

    On Error GoTo ErrHandler
    Const sProcName As String = "add_auxiliary_drive_prices"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "auxiliary_drive_prices"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Preise für Nebentriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): gearbox_prices; T_Preise" & vbLf _
         & "Voraussetzungen: gearbox_prices; T_Preise " & vbLf _
         & "Nachbearbeitung: auxiliary_drive_option_prices" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    ' sSQL = "INSERT INTO " & sTab1 & ( series, `size`, `type`, price ) " _
    '      & "SELECT gearbox_prices.series, gearbox_prices.`size`, T_Preise.KPID, Nz([T_Preise]![Preis1],0) AS Preis " _
    '      & "FROM gearbox_prices INNER JOIN T_Preise ON (gearbox_prices.`size` = T_Preise.BGR) " _
    '      & "AND (gearbox_prices.series = T_Preise.BAR) " _
    '      & "WHERE (((T_Preise.KPID) Like '2A*') AND ((T_Preise.BAF)='V')) " _
    '      & "ORDER BY gearbox_prices.series, gearbox_prices.`size`, T_Preise.KPID ;"

    sSQL = "INSERT INTO " & sTab1 & " ( series, `size`, `type`, price ) " _
         & "SELECT DISTINCTROW gearbox_prices.series, gearbox_prices.`size`, Left([Attr],3) AS NTyp, 0 AS Preis " _
         & "FROM T_Preiskomponenten INNER JOIN (gearbox_prices INNER JOIN T_Preise ON (gearbox_prices.`size` = T_Preise.BGR) " _
         & "AND (gearbox_prices.series = T_Preise.BAR)) ON T_Preiskomponenten.KPID = T_Preise.KPID " _
         & "WHERE (((T_Preise.BAF)='V') AND ((T_Preise.KPID) Like '2[A-D]*')) " _
         & "ORDER BY gearbox_prices.series, gearbox_prices.`size`, Left([Attr],3) ;"

    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere

End Sub

Public Sub add_auxiliary_drive_option_prices(Optional bClean = True)
    ' Purpose:  Add auxiliary drive option prices
    ' Author:   Andreas Herrel
    ' Date:     2018-11-23
    ' Inputs:   gearbox_prices; T_Preise; T_Preiskomponenten
    ' Output:   auxiliary_drive_option_prices
    ' Requires: MySQL database online

    On Error GoTo ErrHandler
    Const sProcName As String = "add_auxiliary_drive_option_prices"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String

    sTab1 = "auxiliary_drive_option_prices"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Preise für Nebentriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): auxiliary_drive_prices; T_Preise" & vbLf _
         & "Voraussetzungen: auxiliary_drive_prices; T_Preise " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Optionen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    sSQL = "INSERT INTO auxiliary_drive_option_prices ( series, [size], type, [option], price ) " _
         & "SELECT DISTINCTROW gearbox_prices.series, gearbox_prices.size, Left([Attr],3) AS NTyp, T_Preise.KPID, Nz([T_Preise]![Preis1],0) AS Preis " _
         & "FROM T_Preiskomponenten INNER JOIN (gearbox_prices INNER JOIN T_Preise ON (gearbox_prices.series = T_Preise.BAR) " _
         & "AND (gearbox_prices.size = T_Preise.BGR)) ON T_Preiskomponenten.KPID = T_Preise.KPID " _
         & "WHERE (((T_Preise.BAF)='V') AND ((T_Preise.KPID) Like '2*')) " _
         & "ORDER BY gearbox_prices.series, gearbox_prices.size, Left([Attr],3) ;"

    Debug.Print sSQL
    DoCmd.RunSQL sSQL

 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere

End Sub

Public Sub add_power_moment_mappings(Optional bClean As Boolean = True)
    ' Purpose:  This procedure deletes existing records in the target table (gearbox_power_moment_mappings) and adds new mappings.
    ' Author:   Andreas Herrel
    ' FirstCreated: 2018-12-17
    ' LastUpdated:  2018-12-17
    ' Inputs:   gearboxes, gearbox_power_moments
    ' Output:   gearbox_power_moment_mappings
    ' Requirements: input tables
    ' Post processing: --

    On Error GoTo ErrHandler
    Const sProcName As String = "add_power_moment_mappings"
    Const RefLevel As Byte = 2
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sClass As String, sFeld1 As String
    
    sTab1 = "gearbox_power_moment_mappings"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Classification Societies ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): gearboxes, gearbox_power_moments " & vbLf _
         & "Voraussetzungen: gearboxes + gearbox_power_moments gefüllt" & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Mappings für Leistungmomente neu einfügen
    DoCmd.SetWarnings False
    
    sSQL = "INSERT INTO " & sTab1 & " ( gearbox_series, gearbox_size, gearbox_design, gearbox_e_motor, ratio, class ) " _
         & "SELECT gearboxes.series, gearboxes.size, gearboxes.design, gearboxes.e_motor, gearbox_power_moments.ratio, gearbox_power_moments.class " _
         & "FROM gearboxes INNER JOIN gearbox_power_moments ON (gearboxes.size = gearbox_power_moments.gearbox_size) " _
         & "AND (gearboxes.series = gearbox_power_moments.gearbox_series);"

    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
    DoCmd.SetWarnings True
    sSQL = "SELECT " & sTab1 & ".* FROM " & sTab1 & " ;"
    Set rs1 = CurrentDb.OpenRecordset(sSQL, dbOpenSnapshot)
    With rs1
        .MoveLast
        sMsg = "Procedure completed normally. " & vbCrLf & CStr(Nz(.RecordCount)) & " records inserted into " & sTab1
        .Close
    End With
    Set rs1 = Nothing
    MsgBox sMsg, vbInformation, sProcName

 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: Modul" & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
End Sub
