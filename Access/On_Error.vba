' Description:
' Author: Andreas Herrel
' Date:
On Error GoTo ErrHandler
    Const sProcName As String = "aaa"
    Dim sKrit as String, sMsg as String, sAns as String, sSQL as String, sFltr As String
    Dim db1 as DAO.Database
    Dim rs1 as DAO.Recordset

ExitHere:
    Exit Sub

ErrHandler:
    With Err
        sMsg = "Object: " & Me.Name & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere