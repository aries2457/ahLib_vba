Public Sub upd_gearbox_prices(Optional bAdd As Boolean = True)
 ' Purpose: Update existing and add new basic prices (KPID = 1G00) for gearboxes
 ' Author:  Andreas Herrel
 ' Date:    2019-04-03;  last updated: 
 ' Inputs:  gearbox_series; T_Preise
 ' Output:  gearbox_prices
 ' Requirements: MySQL database online

     On Error GoTo ErrHandler
     Const sProcName As String = "upd_gearbox_prices"
     Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
     Dim sTab1 as String
     Dim iR1 As Long, iR2 As Long
     Dim db1 As DAO.Database
     Dim rs1 As DAO.Recordset, rs2 As DAO.Recordset

     sTab1 = "gearbox_prices"
     sMsg = "Diese Prozedur aktualisiert vorhandene Einträge in der Zieltabelle " & vbLf _
          & "und fügt neue Preise für Getriebe ein. " & vbLf _
          & "Zieltabelle: " & sTab1 & vbLf _
          & "Quelle(n): CurrentDB.T_Preise_1G00 " & vbLf _
          & "Voraussetzungen: gearbox_series " & vbLf _
          & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen ggf. ergänzt werden." & vbLf & vbLf _
          & "Prozedur fortsetzen?"

     If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
     DoCmd.SetWarnings False

     'Bestehende Preise aktualisieren
     sSQL = "UPDATE " & sTab1 & " INNER JOIN T_Preise_1G00 ON (gearbox_prices.series = T_Preise_1G00.BAR) AND (" & sTab1 & ".size = T_Preise_1G00.BGR) " _
          & "SET " & sTab1 & ".price = (T_Preise_1G00.Preis);"

     Debug.Print sSQL
     DoCmd.RunSQL sSQL
     sSQL = "SELECT gearbox_prices.* " _
          & "FROM gearbox_prices INNER JOIN T_Preise_1G00 ON (gearbox_prices.size = T_Preise_1G00.BGR) AND (gearbox_prices.series = T_Preise_1G00.BAR);"

          Set rs1 = CurrentDB.OpenRecordset(sSQL, dbOpenSnapshot)
     With rs1
          .MoveLast
          iR1 = .RecordCount
          .Close
     End With 
     Set rs1 = Nothing

     'Neue Preise für Grundgetriebe einfügen 
     If bAdd Then
          sSQL = "SELECT T_Preise_1G00.BAR, T_Preise_1G00.BGR, T_Preise_1G00.Preis " _
               & "FROM gearbox_series INNER JOIN (T_Preise_1G00 LEFT JOIN gearbox_prices ON (T_Preise_1G00.BGR = gearbox_prices.size) " _
               & "AND (T_Preise_1G00.BAR = gearbox_prices.series)) ON gearbox_series.name = T_Preise_1G00.BAR " _
               & "WHERE (((gearbox_prices.series) Is Null) AND ((gearbox_prices.size) Is Null));"
          
          Debug.Print sSQL
          Set rs2 = CurrentDB.OpenRecordset(sSQL, dbOpenSnapshot)
          With rs2
               .MoveLast
               iR2 = Nz(.RecordCount)
               .Close
          End With
          Set rs2 = Nothing

          sSQL = "INSERT INTO gearbox_prices ( series, `size`, price ) " _
               & "SELECT T_Preise_1G00.BAR, T_Preise_1G00.BGR, T_Preise_1G00.Preis " _
               & "FROM gearbox_series INNER JOIN (T_Preise_1G00 LEFT JOIN gearbox_prices ON (T_Preise_1G00.BGR = gearbox_prices.size) " _
               & "AND (T_Preise_1G00.BAR = gearbox_prices.series)) ON gearbox_series.name = T_Preise_1G00.BAR " _
               & "WHERE (((gearbox_prices.series) Is Null) AND ((gearbox_prices.size) Is Null));"
          Debug.Print sSQL
          DoCmd.RunSQL sSQL
     
     End If

     sMsg = CStr(iR1) & " Datensätze aktualisiert und" & vbCrLf _
          & CStr(iR2) & " Datensätze neu hinzugefügt"
     MsgBox sMsg, vbInformation, sProcName

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