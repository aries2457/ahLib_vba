Public Sub upd_feature_model()
' Purpose:  Reads file model.xlm and inserts text into table feature_models
' Author:   Andreas Herrel
' Date:     2018-12-18
' Inputs:   model.xlm
' Output:   textstream for feature_models.model
' Requires: Valid XML-file and references for ADO, Scripting, Office libraries 

On Error GoTo ErrHandler
    Const sProcName As String = "upd_feature_model"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 as String, sPNam as String, sFNam as String, sAppl as String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim fso As FileSystemObject
    Dim myStream As Stream

    Set fso = CreateObject("Scripting.FileSystemObject") 
    sTab1 = "feature_models"
    sFNam = "model.xlm"
    sAppl = fso.GetBaseName(CurrentDb.Name)
    sMsg = "Diese Prozedur löscht den vorhandenen Eintrag in der Zieltabelle " & vbLf _
         & "und fügt eine neue Version des Feature Models ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): " & sFNam & vbLf _
         & "Voraussetzungen: valide XML-Datei " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    sPNam = GetSetting(sAppl, "Standard","FeatureModelPath")
    sPNam = IIf(fso.FileExists(sPNam), sPNam, "C:\")
    sPNam = Get1FileDialog(sPNam, "Please select a valid feature model file", "XML")
    If Instr(sPNam, "xml") = 0 Then GoTo ExitHere
    SaveSetting sAppl, "Standard", "FeatureModelPath", sPNam

    Set myStream = New Stream
    With myStream
        .Open
        .LoadFromFile sPNam
        .Charset = "UTF-8"
        .Position = 2

        Set rs1 = CurrentDb.Openrecordset(sTab1, dbOpenDynaset)
        With rs1
            .MoveFirst
            .Edit
            !model = myStream.ReadText(adReadAll)
            .Update
            .Close
        End With
        Set rs1 = Nothing

        .Close
    End With
    Set myStream = Nothing

    sMsg = "Feature model updated succesfully"
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