Public Sub add_gearbox_sap_keys(Optional bClean As Boolean = True)
On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_sap_keys"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset

    sTab0 = "T_Produkthierarchie_VC"
    sTab1 = "gearbox_sap_keys"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue KMATs für den SAP Import ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): gearbox_prices, " & sTab0 & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Datensätze löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    sSQL = "INSERT INTO gearbox_sap_keys ( series, `size`, sap_material_key ) " _
         & "SELECT gearbox_prices.series, gearbox_prices.size, T_Produkthierarchie_VC.KMATNR " _
         & "FROM gearbox_prices INNER JOIN T_Produkthierarchie_VC ON (gearbox_prices.size = T_Produkthierarchie_VC.VC_BAUGROESSE) " _
         & "AND (gearbox_prices.series = T_Produkthierarchie_VC.VC_BAUREIHE) " _
         & "ORDER BY gearbox_prices.series, Val([size]);"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL

    'L-Getriebe werden extra eingefügt
    sSQL = "INSERT INTO gearbox_sap_keys ( series, `size`, sap_material_key ) " _
         & "SELECT DISTINCT gearbox_prices.series, gearbox_prices.size, T_Produkthierarchie_VC.KMATNR " _
         & "FROM gearbox_prices INNER JOIN T_Produkthierarchie_VC ON gearbox_prices.series = T_Produkthierarchie_VC.VC_BAUREIHE " _
         & "WHERE (((gearbox_prices.size) Like '*L') AND ((Val([VC_BAUGROESSE]))=Val([size]))) " _
         & "ORDER BY gearbox_prices.series ;"
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
