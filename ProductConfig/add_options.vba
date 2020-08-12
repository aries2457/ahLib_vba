Public Sub add_options(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_options"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    
    sTab0 = "T_Preiskomponenten"
    sTab1 = "gearbox_options"
    sMsg = "Diese Prozedur l�scht vorhandene Eintr�ge in der Zieltabelle " & vbLf _
         & "und f�gt neue Optionen f�r Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): T_gearbox_options, " & sTab0 & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abh�ngige Tabellen m�ssen neu gef�llt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Datens�tze l�schen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Eintr�ge f�r Optionen aus Preiskomponenten neu einf�gen
    sSQL = "INSERT INTO " & sTab1 & " ( `name` ) " _
         & "SELECT " & sTab0 & ".KPID " _
         & "FROM " & sTab0 & " " _
         & "WHERE " & sTab0 & ".KPID NOT LIKE '0*' " _
         & "ORDER BY " & sTab0 & ".KPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
    'Zus�tzliche Eintr�ge f�r SAP-Import (XML-Datei) aus lokaler Optionstabelle einf�gen
    sTab0 = "T_gearbox_options"
    sSQL = "INSERT INTO " & sTab1 & " ( `name` ) " _
         & "SELECT " & sTab0 & ".( `name` ) " _
         & "FROM " & sTab0 & " ;" 
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