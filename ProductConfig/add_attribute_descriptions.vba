Public Sub add_attribute_descriptions(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_attribute_descriptions"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_attribute_descriptions"
    sMsg = "Diese Prozedur l�scht vorhandene Eintr�ge in der Zieltabelle " & vbLf _
         & "und f�gt neue Bema�ungsattribute f�r Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Attribute" & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abh�ngige Tabellen m�ssen neu gef�llt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Eintr�ge f�r societies l�schen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Eintr�ge f�r societies neu einf�gen (Quelle currentDB)
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