    ' Purpose:  play a game
    ' Author:   Andreas Herrel
    ' Date:     2018-10-17; last updated:
    ' Inputs:   tables and queries
    ' Output:   whatever
    ' Requires: e.g. Reference to object library

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
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere