Private Sub Befehl_Load_Config_XML_Click()
' Purpose:  Load previous saved configuration file with gearbox options
' Author:   Andreas Herrel
' Date:     2018-11-16
' Inputs:   XML text file
' Output:   Marks components in option list
' Requires: FileOpen class modul

On Error GoTo ErrHandler
   Const sProcName As String = "Load_Config_XML"
   Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sLine As String, sFil As String, sOut As String, sTab As String, sTyp As String, sKPID As String
    Dim sField() As String, sValue() As String
    Dim bL As Boolean
    Dim i As Integer, iT As Integer, iK As Integer, iM As Integer, iZ As Integer, n As Integer
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset, rs2 As DAO.Recordset
    Dim fso As FileSystemObject
    Dim tso As Object
    Dim fdo As New FileDialog

    sMsg = "This procedure will load a previously saved option list for a specific gearbox type " _
         & "and will overwrite existing selections for this type." & vbLf & vbLf _
         & "Continue anyway?"
    If MsgBox(sMsg, vbOKCancel + vbExclamation, sProcName) = vbCancel Then Exit Sub
    sOut = GetHyperlinkBase & "Output"
    sTab = "A_Preise0"

    With fdo
        .DialogTitle = "Load configuration"
        .DefaultExt = "XML"
        .DefaultDir = sOut
        .DefaultFileName = Me!TYPID & ".xml"
'        .Flags = OFN_PATHMUSTEXIST
        .MultiSelect = False
        .Filter1Text = "Text Files (.xml)"
        .Filter1Suffix = "*.xml"
        .ShowOpen
        If .FileName = "" Then
            sMsg = "No file selected!"
            MsgBox sMsg, vbCritical, sProcName
            Exit Sub
        End If
        sFil = .FileName
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    ImportXML DataSource:=sFil, ImportOptions:=acStructureAndData
    
Exit Sub 

    sTyp = sValue(iT)
    sSQL = "UPDATE " & sTab & " SET " & sTab & ".Sel = False " _
         & "WHERE (" & sTab & ".TYPID = '" & sTyp & "') ;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL sSQL

    sSQL = "SELECT " & sTab & ".* FROM " & sTab & " " _
         & "WHERE (" & sTab & ".TYPID = '" & sTyp & "') " _
         & "ORDER BY " & sTab & ".KPID ;"
    Set rs1 = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset)

    With rs1
        .FindFirst "[KPID]='" & sValue(iK) & "' AND [Knz]='" & sValue(iZ) & "'"
        bL = IIf(InStr(Nz(sValue(iZ)), "L") <> 0, True, False)
        If Not .NoMatch Then
            .Edit
            !Sel = True
            !Mg = Val(Nz(sValue(iM)))
            .Update
        End If
        .MoveFirst
        i = 1

        Do While tso.AtEndOfStream <> True
            sValue = Split(tso.ReadLine, ";")
            .FindFirst "[KPID]='" & sValue(iK) & "' AND [Knz]='" & sValue(iZ) & "'"
            If InStr(Nz(sValue(iZ)), "L") <> 0 Then bL = True
            If Not .NoMatch Then
                .Edit
                !Sel = True
                !Mg = Val(Nz(sValue(iM)))
                .Update
            End If
            Debug.Print sTyp, sValue(iK)
            .MoveFirst
            i = i + 1
            If i > 100 Then Exit Do
        Loop

    End With
    
    Debug.Print "DONE", i
    tso.Close
    
    Set tso = Nothing
    Set fso = Nothing

    Set rs2 = Me.RecordsetClone
    With rs2
        .FindFirst "[TYPID]='" & sTyp & "'"
        Me.Bookmark = .Bookmark
        .Close
    End With
    Set rs2 = Nothing
    Me!Kontroll_L = bL
    Call Filter4Options
    
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