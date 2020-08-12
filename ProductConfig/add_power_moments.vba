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