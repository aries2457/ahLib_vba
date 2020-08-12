Private Sub Befehl_Save_Config_Click()
' Purpose:  Save selected gearbox options into configuration file
' Author:   Andreas Herrel
' Date:     2018-10-26
' Inputs:   A_Preise0
' Output:   CSV text file 
' Requires: FileDialog class module

On Error GoTo ErrHandler
    Const sProcName As String = "Save_Config"
    Const ForWriting As Long = 2
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sLine As String, sFil As String, sOut As String, sTab As String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim fso As FileSystemObject
    Dim tso As Object
    Dim fdo As New FileDialog

    sOut = GetHyperlinkBase & "Output"
    With fdo
        .DialogTitle = "Save configuration"
        .DefaultExt = "CSV"
        .DefaultDir = sOut
        .DefaultFileName = Me!TYPID & ".csv"
'        .Flags = OFN_PATHMUSTEXIST
        .MultiSelect = False
        .Filter1Text = "Text Files (.csv)"
        .Filter1Suffix = "*.csv"
        .ShowSave
        If .FileName = "" Then Exit Sub
        sFil = .FileName
    End With
    
    If sFil = "" Then
        sMsg = "No file selected!"
        MsgBox sMsg, vbCritical, sProcName
        Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tso = fso.CreateTextFile(sFil, True)

    sTab = "A_Preise0"
    sSQL = "SELECT " & sTab & ".* FROM " & sTab & " " _
         & "WHERE (" & sTab & ".TYPID = '" & Me!TYPID & "' AND " & sTab & ".Sel = True) " _
         & "ORDER BY " & sTab & ".KPID ;"
    Set rs1 = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset)
    With rs1
        If .EOF Then 
            MsgBox "Nothing selected!", vbCritical, sProcName
        Else
            sLine = "BAR;BGR;Knz;BAF;MOT;KPID;Mg;Bez_en;TYPID;Knz"
            tso.WriteLine (sLine)
        End If
        Do While Not .EOF
            sLine = !BAR & ";" & !BGR & ";" & !Knz & ";" & !BAF & ";" & !MOT & ";" & !KPID & ";" & !Mg & ";" & !Bez_en & ";" & !TYPID & ";" & !Knz
            tso.WriteLine (sLine)
            .MoveNext
        Loop
        .Close
    End With
    Set rs1 = Nothing
    
    tso.Close
    Set tso = Nothing

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
