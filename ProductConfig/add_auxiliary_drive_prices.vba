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