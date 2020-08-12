Public Sub add_gearbox_prices(Optional bClean As Boolean = True)
 ' Purpose: Add basic prices for gearboxes
 ' Author:  Andreas Herrel
 ' Date:    2019-04-02;  last updated: 
 ' Inputs:  gearbox_series; T_Preise
 ' Output:  gearbox_prices
 ' Requirements: MySQL database online

    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_prices"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 as String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset

    sTab1 = "gearbox_prices"
    sMsg = "Diese Prozedur l�scht vorhandene Eintr�ge in der Zieltabelle " & vbLf _
         & "und f�gt neue Preise f�r Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_Preise " & vbLf _
         & "Voraussetzungen: gearbox_series " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abh�ngige Tabellen m�ssen neu gef�llt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    DoCmd.SetWarnings False
    
    'Bestehende Eintr�ge f�r Baureihen l�schen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Neue Preise f�r Grundgetriebe einf�gen 
    sSQL = "INSERT INTO gearbox_prices ( `series`, `size`, `price` ) " _
         & "SELECT T_Preise.BAR, T_Preise.BGR, Nz(`Preis1`,0) AS Preis " _
         & "FROM gearbox_series INNER JOIN T_Preise ON gearbox_series.name = T_Preise.BAR " _
         & "WHERE ((T_Preise.KPID)='1G00') " _
         & "ORDER BY T_Preise.BAR, T_Preise.BGR ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
   Exit Sub

 ErrHandler:
   With Err
      sMsg = "Object: " & Me.Name & vbCrLf _
           & "Procedure: " & sProcName & vbCrLf _
           & "Error: " & .Number & vbCrLf & .Description
      Debug.Print sMsg
      MsgBox sMsg, vbCritical + vbOKOnly, .Source
   End With
   Resume ExitHere

End Sub