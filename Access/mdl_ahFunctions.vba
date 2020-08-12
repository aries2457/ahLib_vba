Option Compare Database
Option Explicit

Public Const AppName As String = "Product Guide Basic"
Public Const AppVers As String = "0.8.3"

Public Const conDynaset = 0
Public Const conSnapshot = 2


Private Declare Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function getOSUserName() As String
     Dim lngLen As Long, lngX As Long
     Dim strUserName As String
    
     strUserName = String$(254, 0)
     lngLen = 255
     lngX = apiGetUserName(strUserName, lngLen)
     If lngX <> 0 Then
        getOSUserName = Left$(strUserName, lngLen - 1)
       Else
        getOSUserName = ""
     End If
End Function

Public Function getDomain()
    Dim objWMISvc As Object, colItems As Object, objItem As Object
    Dim strComputerDomain As String
    
    Set objWMISvc = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMISvc.ExecQuery("SELECT * FROM Win32_ComputerSystem", , 48)
    For Each objItem In colItems
        strComputerDomain = objItem.Domain
        If objItem.PartOfDomain Then
            'Debug.Print "Computer Domain: " & strComputerDomain
            getDomain = strComputerDomain
          Else
            'Debug.Print "Workgroup: " & strComputerDomain
            getDomain = strComputerDomain
        End If
        
    Next objItem
    
End Function

Public Function Wait(MilliSekunden As Double)
    Dim i As Double, Ende As Double

    Ende = Timer + (MilliSekunden / 1000)
    Do While i < Ende
        DoEvents
        i = Timer
    Loop
End Function

Function AddAppProperty(strName As String, _
        varType As Variant, varValue As Variant) As Integer
    Dim dbs As Object, prp As Variant
    Const conPropNotFoundError = 3270

    Set dbs = CurrentDb
    On Error GoTo AddProp_Err
    dbs.Properties(strName) = varValue
    AddAppProperty = True

AddProp_Bye:
    Exit Function

AddProp_Err:
    If Err = conPropNotFoundError Then
        Set prp = dbs.CreateProperty(strName, varType, varValue)
        dbs.Properties.Append prp
        Resume
    Else
        AddAppProperty = False
        Resume AddProp_Bye
    End If
End Function

Public Function InitApp(Optional bRefAllways As Boolean = False)
On Error GoTo ErrHandler
    'App soll nur in bestimmten Domänen unter Access Vollversion laufen, sonst nur Runtime
    Const sProcName As String = "InitApp"
    Const DB_Text As Long = 10
    Dim sKrit As String, sMsg As String, sAns As String, sSql As String
    Dim dbs As DAO.Database
    Dim tdf As TableDef
    Dim intX As Integer

    'Eingebundene Tabellen neu verknüpfen, wenn CurrentProject.Path abweicht
    Set dbs = CurrentDb
    
    For Each tdf In dbs.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If InStr(tdf.Connect, CurrentProject.Path) > 0 And Not bRefAllways Then
                Exit For
            Else
                tdf.Connect = "MS Access;PWD=init123;DATABASE=" & CurrentProject.Path & "\" & sBackEnd
                tdf.RefreshLink
            End If
        End If
    Next tdf
    
    'User Interface anpassen, DB-Properties setzen
    AddAppProperty "AppName", DB_Text, AppName
    AddAppProperty "AppVers", DB_Text, AppVers
    AddAppProperty "AppTitle", DB_Text, AppName & "  " & AppVers
    AddAppProperty "AppIcon", DB_Text, CurrentProject.Path & sGraphics & "\LMom32.ico"
    CurrentDb.Properties("UseAppIconForFrmRpt") = 1
    Application.RefreshTitleBar

    'Objekte ausblenden, wenn App in fremden Domains läuft
    If InStr("reintjes.loc,HERREL", getDomain()) > 0 Then
        EnableShift True
        hideObjects "sTables", "T_", False
        hideObjects "sQueries", "A_", False
        SaveSetting AppName, "Standard", "DomType", "intern"
    Else
        EnableShift False
        hideObjects "sTables", "T_", False
        hideObjects "sQueries", "A_", False
        SaveSetting AppName, "Standard", "DomType", "extern"
        
        'App beenden, wenn nicht im Runtime Mode gestartet
        If SysCmd(acSysCmdRuntime) = False Or Application.UserControl = False Then
            MsgBox "Please use ACCESS runtime version to run this application", vbCritical
            Application.Quit
        End If
            
    End If
    
ExitHere:
    Exit Function

ErrHandler:
    With Err
        sMsg = "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
       
End Function

