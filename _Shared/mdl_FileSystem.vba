Option Compare Database
Option Explicit

Public Function GetFolderDialog(Optional sInitialPath = "C:\")
' Purpose: Lässt den Benutzer einen Ordner im Dateisystem auswählen
' Author: Andreas Herrel
' Date:
' Requires reference to Microsoft Office Object Library
On Error GoTo ErrHandler
    Const sProcName As String = "GetFolderDialog"
    Dim sKrit as String, sMsg as String, sAns as String, sSQL as String, sFltr As String
    Dim sFolder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = sInitialPath
        .Title = "Select path to Front-End"
        .ButtonName = "Select Folder"
        .InitialView = msoFileDialogViewList
        If .Show = -1 Then
            sFolder = .SelectedItems(1)
          Else
            sFolder = ""
        End If
    End With
    
    GetFolderDialog = sFolder
ExitHere:
    Exit Function

ErrHandler:
    With Err
        sMsg = "Object: Modul mdl_FileSystem" & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere   
End Function

Public Function GetFileDialog() 
' Purpose: Select one or more files from file system
' Author:  Andreas Herrel
' Date:
' Inputs:
' Output:  Name of one or more files
' Requirements: Reference to Microsoft Office Object Library

On Error GoTo ErrHandler
    Const sProcName As String = "GetFileDialog"
    Dim sKrit as String, sMsg as String, sAns as String, sSQL as String, sFltr As String

    Dim fDialog As Office.FileDialog 
    Dim varFile As Variant 

    ' Clear listbox contents. 
    Me.FileList.RowSource = "" 

    ' Set up the File Dialog. 
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker) 

    With fDialog 

        ' Allow user to make multiple selections in dialog box 
        .AllowMultiSelect = True 

        ' Set the title of the dialog box. 
        .Title = "Please select one or more files" 

        ' Clear out the current filters, and add our own. 
        .Filters.Clear 
        .Filters.Add "Access Databases", "*.ACCDB" 
        .Filters.Add "All Files", "*.*" 

        ' Show the dialog box. If the .Show method returns True, the 
        ' user picked at least one file. If the .Show method returns 
        ' False, the user clicked Cancel. 
        If .Show = True Then 

            'Loop through each file selected and add it to our list box. 
            For Each varFile In .SelectedItems 
                Me.FileList.AddItem varFile 
            Next 

        Else 
            MsgBox "You clicked Cancel in the file dialog box." 
        End If 
    
    End With 

ExitHere:
    Exit Function

ErrHandler:
    With Err
        sMsg = "Object: " & Me.Name & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere

End Function