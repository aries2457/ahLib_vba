Sub AlterTableX(Optional bRestore As Boolean = False)
    ' Purpose:  Fügt mit Hilfe der SQL Data Definition Language (DDL) Felder und Foreign Keys
    '           zu einer angegebenen Tabelle in einer externen DB hinzu.
    ' Author:   Andreas Herrel
    ' Date:     2019-01-30
    ' Inputs:   Original Database, Tables
    ' Output:   Changed Database, Tables
    ' Requires: Exclusive access to target DB
    '           Connection to target DB via ADO, because DAO does not support CASCADE Statements für Referenzielle Integrität
    '           Reference to object ADO libraries
    ' ATTENTION: Leads to runtime error, if fields an relations in target DB already exists

    On Error GoTo ErrHandler
    Const sProcName As String = "AlterTableX"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sDat(2) As String, sConn As String
    Dim oConn As ADODB.Connection
    Dim fso As FileSystemObject

    sDat(0) = "O:\Bereiche_oeffentlich\Bereich_T\Access\Kundenprojekte\Daten\SalesProjects-BE.accdb"  'released Version
    sDat(1) = "M:\TN\Normung\Anwendungsdaten\Access2010\Kundenprojekte\Daten\SalesProjects-BE.accdb"  'test Version
    
    Set oConn = New ADODB.Connection
    sConn = "Provider=Microsoft.ACE.OLEDB.12.0; " & _
            "Data Source=" & sDat(1) & ";" & _
            "User Id=Admin; " & _
            "Mode=Share Deny None; " & _
            "Persist Security Info=False"
    oConn.Open sConn
    sMsg = IIf(bRestore, "Start processing and release ...", "Start processing for testing ...")
    Debug.Print sProcName & ":", sMsg, Now
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print "executing:", "CopyFile to test environment"
    fso.CopyFile sDat(0), sDat(1)

    With oConn
'        GoTo CopyFiles
    
Projekte:
        sSQL = "DROP INDEX PrimaryKey ON T_Projekte_V ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "ALTER TABLE T_Projekte_V ADD PRIMARY KEY (PID) ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
Versionen:
        sSQL = "DROP INDEX PrimaryKey ON T_Versionen ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "ALTER TABLE T_Versionen ADD COLUMN PID Long ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "ALTER TABLE T_Versionen ADD PRIMARY KEY (VID) ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "CREATE INDEX PIDx ON T_Versionen (PID) ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "UPDATE T_Projekte_V INNER JOIN T_Versionen ON (T_Projekte_V.UPNr = T_Versionen.UPNr) " _
             & "AND (T_Projekte_V.PNr = T_Versionen.PNr) AND (T_Projekte_V.PJahr = T_Versionen.PJahr) " _
             & "SET T_Versionen.PID = [T_Projekte_V]![PID];"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "DELETE T_Versionen.* From T_Versionen WHERE (((T_Versionen.PID) Is Null)); "
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "ALTER TABLE T_Versionen ADD CONSTRAINT FK_T_Versionen " _
             & "FOREIGN KEY (PID) REFERENCES T_Projekte_V (PID) " _
             & "ON UPDATE CASCADE ON DELETE CASCADE ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL

Nebentriebe:
        sSQL = "DROP INDEX PrimaryKey ON T_Nebentriebe ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "ALTER TABLE T_Nebentriebe ADD PRIMARY KEY (NID) ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "DELETE T_Nebentriebe.* From T_Nebentriebe WHERE (((T_Nebentriebe.PJahr) < 2000)); "
        Debug.Print "executing: ", sSQL
        .Execute sSQL

        sSQL = "ALTER TABLE T_Nebentriebe ADD COLUMN VID Long ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "CREATE INDEX VIDx ON T_Nebentriebe (VID) ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "UPDATE T_Versionen INNER JOIN T_Nebentriebe ON (T_Versionen.Vers = T_Nebentriebe.Vers) " _
             & "AND (T_Versionen.UPNr = T_Nebentriebe.UPNr) AND (T_Versionen.PNr = T_Nebentriebe.PNr) " _
             & "AND (T_Versionen.PJahr = T_Nebentriebe.PJahr) SET T_Nebentriebe.VID = [T_Versionen]![VID];"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "DELETE T_Nebentriebe.* From T_Nebentriebe WHERE (((T_Nebentriebe.VID) Is Null)); "
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        sSQL = "ALTER TABLE T_Nebentriebe ADD CONSTRAINT FK_T_Nebentriebe " _
             & "FOREIGN KEY (VID) REFERENCES T_Versionen (VID) " _
             & "ON UPDATE CASCADE ON DELETE CASCADE ;"
        Debug.Print "executing: ", sSQL
        .Execute sSQL
        
        .Close
    End With

CopyFiles:
    If bRestore Then
        Debug.Print "executing:", "CopyFile to origin"
        sDat(2) = fso.GetParentFolderName(sDat(0)) & "\" & fso.GetBaseName(sDat(0)) & "_" _
                & Format(Now(), "yyyymmdd_hhnn") & "." & fso.GetExtensionName(sDat(0))
'        Debug.Print sDat(2)
        fso.CopyFile sDat(0), sDat(2)  'Backup original file
        fso.CopyFile sDat(1), sDat(0)  'Restore changed file
    End If

    Set fso = Nothing
    sMsg = "Procedure finished"
    Debug.Print sProcName & ":", sMsg, Now
 
ExitHere:
    Exit Sub

ErrHandler:
    With Err
        sMsg = "Object: " & "Modul" & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
       
End Sub

