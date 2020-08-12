' Attribute VB_Name = "mdlINIfile"
' INI-Files
'  This module contains functions to read from and write
'  to an INI-file. Each project can have just one INI-file
'  and it is located in a subfolder of the default Windows
'  AppData folder (for example C:\Users\USERNAME\AppData\Roaming). The name of the subfolder is the
'  application title (see GetAppicationTitle). Suppose you have
'  a project named VBADocSys, than the INI file can be C:\Users\USERNAME\AppData\Roaming\VBADocSys\VBADocSys.ini.
' ' Save the last access date in section "access"
' INI_SetValue( "access.last", Format( now(), "yyyy-mm-dd" )
' ' Perform installation actions
' If INI_GetValue( "config.version" ) = "2.1" Then
'    ..... ' Some version specific code here
'    INI_SetValue( "config.version", "2.2" )
' End If
' 
'  Content of INI-file:
' [access]
' last=2016-07-12
' [config]
' version=2.2
' autosave=1
' ...
' GetApplicationFolder
' It is recommended by Microsoft to store user-settings in the Windows-registry. But, some companies
' follow a very strict authorization policy and do not allow all users to change the registry. In these
' cases I use an INI-file to store user-settings.
Option Compare Database
Option Explicit
'
'   Maximum buffer size
'
Private Const MAX_BUF_SIZE = 1024
'
'   Default sectionname
'
Private Const cDefaultSection = "config"
'
' Windows API to read from an INI-file.
' section Name of the section in the INI-file
' key     Name of the key within the section speficied
' default Value to be returned when key does not exist
' returnValue Reference to buffer to store the retrieved value
' nSize   The maximum length of the buffer to store the retrieved value
' filename Full path to the INI-file
' Length of the returned value
' INI_GetValue
Private Declare Function GetPrivateProfileString _
      Lib "kernel32" Alias "GetPrivateProfileStringA" _
     (ByVal section As String, _
      ByVal key As Any, _
      ByVal default As String, _
      ByVal returnValue As String, _
      ByVal nSize As Long, _
      ByVal fileName As String) As Long
' Windows API to write to an INI-file.
' section Name of the section in the INI-file
' key     Name of the key within the section speficied
' value   The value to be assigned to the key
' filename Full path to the INI-file
' A non-zero value if the call was succesfull, otherwise zero.
' f the INI-file does not exist it will be created
'  INI_SetValue
Private Declare Function WritePrivateProfileString _
      Lib "kernel32" Alias "WritePrivateProfileStringA" _
     (ByVal section As String, _
      ByVal key As Any, _
      ByVal Value As Any, _
      ByVal fileName As String) As Long
' Returns a @Boolean indication if the item specified by index
' exists.
' index The index of the key in the INI-file (section.key)
' if the key exists within the given section, @False otherwise
' NI_GetValue
' ' Check INI file for the key "version" in section "database"
' If INI_KeyExist( "database.version" ) Then
'    Debug.Print "Database initialized!"
' Else
'    ShowError "Database not connected!"
' End If
Public Function INI_KeyExists(ByVal Index As String) As Boolean
    Dim default  As String
    ' Create a time-dependend default value for the item
    default = Format$(Now(), "yyyymmddHhNnSs")
    
    ' If default is returned, the key does not exist
    If INI_GetValue(Index, default) = default Then
        INI_KeyExists = False
    Else
        INI_KeyExists = True
    End If
End Function
' Returns a @Boolean indication if the section exists
' section The name of the section to be searched
' if the section exists, @False if not
' NI_GetSections
Public Function INI_SectionExists(ByVal section As String) As Boolean
    Dim vSection
    Dim colSections As Collection
    ' Get list of all sections in the INI-file
    Set colSections = INI_GetSections()
    ' Search for the section specified
    For Each vSection In colSections
        If vSection = section Then
            INI_SectionExists = True
            Exit Function
        End If
    Next vSection
    ' When arrivered over here, the section has not been found
    INI_SectionExists = False
End Function
' Returns a @String containing the value of the specified item
' index The index of the key in the INI-file (section.key)
' default The value to be returned if the key does not exist
' The value of the key. If the key does not exists the default value is returned
' etPrivateProfileString
Public Function INI_GetValue(ByVal Index As String, _
                             Optional ByVal default As String = "") As String
    Dim buffer  As String
    Dim length  As Long
    
    buffer = Space$(MAX_BUF_SIZE)
    length = GetPrivateProfileString( _
                GetSectionFromIndex(Index), _
                GetKeyFromIndex(Index), _
                default, _
                buffer, _
                Len(buffer), _
                INI_Filename())
    ' Skip trailing characters of buffer
    INI_GetValue = Left$(buffer, length)
End Function
' Assigns a value to a specified key
' index The index of the key in the INI-file (section.key)
' value The value to be assignd to the key
' WritePrivateProfileString
Public Sub INI_SetValue(ByVal Index As String, Value As String)
    Call WritePrivateProfileString( _
             GetSectionFromIndex(Index), _
             GetKeyFromIndex(Index), _
             Value, _
             INI_Filename())
End Sub
' Returns @Dictionary containing all keys in the specified section and their values
' Dictionary
' section Name of the section to be read
' The keys correspond to the keys within the section
' im dict As Dictionary
' Dim key
' Set dict = INI_GetSection("config")
' For Each k In dict.Keys
'     Debug.Print key, dict(key)
' Next k
' GetPrivateProfileString
' INI_GetValue
Public Function INI_GetSection(ByVal section As String) As Dictionary
    Dim buffer  As String
    Dim length  As Long
    Dim arrKeys
    Dim result  As New Dictionary
    Dim i       As Long
    '   Create buffer (large)
    buffer = Space$(MAX_BUF_SIZE * 10)
    '   Get list of all keys within section (key = vbNullChar)
    length = GetPrivateProfileString( _
                section, _
                vbNullString, _
                "", _
                buffer, _
                Len(buffer), _
                INI_Filename())
    '   The keys are separated by Null-values
    arrKeys = Split(Left$(buffer, length), vbNullChar)
    '   Last item is also followed by a Null-value
    For i = LBound(arrKeys) To UBound(arrKeys) - 1
        ' Retrieve key value and store it in a Dictionary
        result.Add arrKeys(i), INI_GetValue(section & "." & arrKeys(i))
    Next i
    
    Set INI_GetSection = result

End Function
' Replaces or inserts a new section in the INI-file and adds the content
' of values to this section.
' section Name of the section to be added or inserted
' values  Content of the section. It should contain at least one element.
' When values has no elements no section will be inserted. If the
' section already exists, it will be deleted.
' Dictionary
' INI_SectionExists
' INI_Remove
' INI_SetValue
Public Sub INI_SetSection(ByVal section As String, values As Dictionary)
    Dim varKey
    '  Section exists? Remove it!
    If INI_SectionExists(section) Then
        Call INI_Remove(section)
    End If
    '
    For Each varKey In values.Keys
        Call INI_SetValue(section & "." & varKey, values.Item(varKey))
    Next varKey
End Sub
' Returns a @Collection of sections within the INI-file
' List of sections.
' etPrivateProfileString
' ArrayToCollection
Public Function INI_GetSections() As Collection
    Dim buffer  As String
    Dim length  As Long
    Dim sections
    '   Maak buffers
    buffer = Space$(MAX_BUF_SIZE)
    length = GetPrivateProfileString( _
                  vbNullString, _
                  vbNullString, _
                  "", _
                  buffer, _
                  Len(buffer), _
                  INI_Filename())
    '   sections are separated by a Null-character
    sections = Split(Left$(buffer, length), vbNullChar)
    ' The last section is also followed by a Null-char!
    ' So shorten the array by 1
    Set INI_GetSections = ArrayToCollection(sections, endIndex:=UBound(sections) - 1)
End Function
' Removes all sections from the INI-file
' INI_GetSections
' INI_Remove
Public Sub INI_RemoveAll()
    Dim varSectie
    Dim colSecties      As Collection
    
    Set colSecties = INI_GetSections()
    For Each varSectie In colSecties
        Call INI_Remove(varSectie)
    Next varSectie

End Sub
' Remove a section or key from the INI-file
' index Pair of section and key, separated by a dot or only the name of the section to be removed.
' INI_SectionExists
' INI_KeyExists
' WritePrivateProfileString
' When the INI-file contains a section "database" and the default section also contains a key
' "database" then INI_Remove( "database" ) will just remove the section "database". The key
' "config.database" will remain! If you want to remove this key, call INI_Remove( "config.database" )
Public Sub INI_Remove(ByVal Index As String)
    ' Remove section if index is a section
    If INI_SectionExists(Index) Then
        Call WritePrivateProfileString(Index, vbNullString, vbNullString, INI_Filename())
    ElseIf INI_KeyExists(Index) Then
        ' Remove key
        Call WritePrivateProfileString( _
                 GetSectionFromIndex(Index), _
                 GetKeyFromIndex(Index), _
                 vbNullString, _
                 INI_Filename())
    End If
    
End Sub
' @Private function which returns a @String containing the name of the
' section. This is the value before the first dot in index.
' When no section is specified, the default section ('config') is
' returned.
' Name of the section
' ndex Pair of section and key, separated by a dot.
Private Function GetSectionFromIndex(Index As String) As String

    Dim pos As Integer
    
    pos = InStr(Index, ".")
    If pos = 0 Then
        GetSectionFromIndex = cDefaultSection
    Else
        GetSectionFromIndex = Left$(Index, pos - 1)
    End If
    
End Function
' @Private function which returns a @String containing the name of the
' key. This is the value after the first dot in index.
' Name of the key
' ndex Pair of section and key, separated by a dot.
Private Function GetKeyFromIndex(Index As String) As String
    Dim pos As Integer
    
    pos = InStr(Index, ".")
    If pos = 0 Then
        GetKeyFromIndex = Index
    Else
        GetKeyFromIndex = Mid$(Index, pos + 1)
    End If
End Function
' This @Private function returns the full path of the
' INI-file of the current project.
' Path to the INI-file.
' he file name is determined once and stored as a @static
' variable. If you set the AppTitle property after this
' function is called, the name of the INI-file will not
' change.
' When the AppTitle property is set, the first word of this property
' is the name of the application. I am used to set this property
' as follows: {application} V{version}. I do not want the version
' number in the name of the INI-file
' Filesystem
' GetApplicationFolder
' GetApplicationTitle
' GetWord
Private Function INI_Filename() As String

    Static fileName     As String
    
    If IsEmptyVar(fileName) Then
        fileName = GetApplicationFolder() & "\" & GetWord(GetApplicationTitle(), 1) & ".ini"
    End If
    
    INI_Filename = fileName
End Function
