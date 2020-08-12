Private Sub Befehl_saveConfig_XML_Click()

On Error GoTo ErrHandler
    Const sProcName As String = "saveConfig_XML"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sQuery As String, sWhere As String, sOut As String, sFil As String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim fso As FileSystemObject
    Dim fdo As New FileDialog
    

    sOut = GetHyperlinkBase & "Output"
    With fdo
        .DialogTitle = "Save configuration"
        .DefaultExt = "XML"
        .DefaultDir = sOut
        .DefaultFileName = Me!TYPID & ".xml"
'        .Flags = OFN_PATHMUSTEXIST
        .MultiSelect = False
        .Filter1Text = "Text Files (.xml)"
        .Filter1Suffix = "*.xml"
        .ShowSave
        If .FileName = "" Then Exit Sub
        sFil = .FileName
    End With
    
    If sFil = "" Then
        sMsg = "No file selected!"
        MsgBox sMsg, vbCritical, sProcName
        Exit Sub
    End If

    sQuery = "A_Preise0"
    sWhere = sQuery & ".TYPID = '" & Me!TYPID & "' AND " & sQuery & ".Sel = True"
    ExportXML acExportQuery, sQuery, sFil, WhereCondition:=sWhere

ExitHere:
    Exit Sub

ErrHandler:
    With Err
        sMsg = "Object: " & "Modul" & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere

End Sub
