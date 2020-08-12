VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageHeaderClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
' * This class load a TIF(Group IV), JFIF(JPEG) and PNG image header and
' * return the Width, Height, Resolution, BitsperPixel and Grayscale mode.
' *
' * For more details about other images formats visit http://www.wotsit.org
' *
' * Paulo S. Pedrozo - 25/06/2003
' * Santa Catarina - Brasil
' ************************************************************************

Option Explicit

Private FileNumber As Byte
Private Buffer() As Byte

Private Type TypeImgHeader
    Height As Integer
    Width As Integer
    DPI As Integer
    BitsPerPixel As Byte
    IsGray As Boolean
End Type
Private ImgHeader As TypeImgHeader
Function Load(Filename As String) As Boolean
        
    ' *******************************************************************
    ' * Detect the image file type and call the respective routine
    ' *******************************************************************
    
    Dim CodRet As Boolean
    
    On Error Resume Next
    
    If Dir(Filename) = "" Then Exit Function
    
    ' Reset the variables
    ImgHeader.Height = 0
    ImgHeader.BitsPerPixel = 0
    ImgHeader.Width = 0
    ImgHeader.DPI = 0
    ImgHeader.IsGray = False
    
    ' Open the file
    FileNumber = FreeFile()
    Open Filename For Binary As #FileNumber
    If Err <> 0 Then Exit Function
    If LOF(FileNumber) = 0 Then Exit Function
    
    ' Read the first 2 bytes to identifier
    ReDim Buffer(0 To 1)
    Get #FileNumber, , Buffer
    
    ' Tiff
    If Buffer(0) = &H49 And Buffer(1) = &H49 Then
        CodRet = GetHeaderTIF
    End If

    ' Jpeg
    If Buffer(0) = &HFF And Buffer(1) = &HD8 Then
        CodRet = GetHeaderJPG
    End If
    
    ' PNG
    If Buffer(0) = &H89 And Buffer(1) = &H50 Then
        CodRet = GetHeaderPNG
    End If
        
    ' Close the file
    Close #FileNumber
        
    If CodRet Then Load = True

End Function
Private Function GetHeaderPNG() As Boolean

    ' ********************************************************************
    ' * The PNG format work with Chunk, each Chunk have variable size and
    ' * have some informations, in this class I just use the IHDR and pHYs
    ' * The IDAT and IEND chunk is the begin and end of image data.
    ' * For more details visit http://www.libpng.org/pub/png/spec/PNG-Chunks.html
    ' ********************************************************************
    
    Dim ChunkSize As Long
    Dim ChunkType As String
    
    ' Get the next 6 bytes to confirms that a PNG
    ReDim Buffer(0 To 5)
    Get #FileNumber, , Buffer
    If Buffer(0) <> &H4E Or Buffer(1) <> &H47 Or Buffer(2) <> &HD Or _
       Buffer(3) <> &HA Or Buffer(4) <> &H1A Or Buffer(5) <> &HA Then
        Exit Function
    End If
    
    ' Dim the array. Both the chunk size and chunk type have 4 bytes
    ReDim Buffer(0 To 3)
        
    Do While True
        ' Get the chunk size
        Get #FileNumber, , Buffer
        ChunkSize = (Buffer(0) * (2 ^ 24)) + (Buffer(1) * (2 ^ 16)) + (Buffer(2) * (2 ^ 8)) + Buffer(3)
        
        ' Get the chunk type
        Get #FileNumber, , Buffer
        ChunkType = Chr(Buffer(0)) & Chr(Buffer(1)) & Chr(Buffer(2)) & Chr(Buffer(3))
        
        ' Set the array to chunk size all chunk
        ReDim Buffer(0 To ChunkSize - 1)
        Get #FileNumber, , Buffer
        
        Select Case ChunkType
            Case "IHDR"
                ' Main Chunk with Height, Width and BitsPerPixel
                ImgHeader.Width = (Buffer(0) * (2 ^ 24)) + (Buffer(1) * (2 ^ 16)) + (Buffer(2) * (2 ^ 8)) + Buffer(3)
                ImgHeader.Height = (Buffer(4) * (2 ^ 24)) + (Buffer(5) * (2 ^ 16)) + (Buffer(6) * (2 ^ 8)) + Buffer(7)
                ImgHeader.BitsPerPixel = Buffer(8)
                    
                ' Composition to found the correct bitsperpixel
                If ImgHeader.BitsPerPixel = 8 Then
                    If Buffer(9) = 0 Or Buffer(9) = 4 Then ImgHeader.IsGray = True
                    If Buffer(9) = 2 Or Buffer(9) = 6 Then ImgHeader.BitsPerPixel = 24
                End If
                
                        
            Case "pHYs"
                ' Resolution
                If Buffer(8) = 1 Then
                    ImgHeader.DPI = Round(((Buffer(0) * (2 ^ 24)) + (Buffer(1) * (2 ^ 16)) + (Buffer(2) * (2 ^ 8)) + Buffer(3)) * 0.0254)
                End If
                
            Case "IDAT", "IEND"
                ' End
                Exit Do
                
        End Select
    
        ReDim Buffer(0 To 3)
        
        ' Jump 4 bytes
        Get #FileNumber, , Buffer
    Loop
    
    GetHeaderPNG = True
    
End Function

Private Function GetHeaderTIF() As Boolean

    ' *******************************************************************
    ' * The TIFF group IV work with many Information File Directory(IFD)
    ' * each IFD have 12 bytes the first 2 bytes is the TAG name
    ' *
    ' * The others IFDs visit www.wotsit.com
    ' *
    ' *******************************************************************
    
    Dim BufferDPI() As Byte
    Dim IDTag As Long
    Dim Pos As Long, PosIFD As Long
    Dim Offset As Long
    Dim TotalIFD, F As Integer
    Dim Resolution As Double
    Dim ResolutionUnit As Byte

    ' Get the rest of fix initial header
    ReDim Buffer(0 To 5)
    Get #FileNumber, , Buffer
    
    ' Only this TIFF type is supported
    If Buffer(0) <> 42 Then Exit Function
    
    ' Locate where is the IFD totals information
    PosIFD = (Buffer(2) + Buffer(3) * (2 ^ 8) + Buffer(4) * (2 ^ 16) + (Buffer(5) * (2 ^ 24))) + 1
    
    ' get the numbers of IFD
    ReDim Buffer(0 To 1)
    Get #FileNumber, PosIFD, Buffer
    TotalIFD = Buffer(0) + (Buffer(1) * (2 ^ 8))
    
    ' Set the buffer to IFD size(12 Bytes)
    ReDim Buffer(0 To (TotalIFD * 12) - 1)
    Get #FileNumber, PosIFD + 2, Buffer
        
    ' Loop each IFD and use only Height, Width and Resolution
    Pos = 0
    For F = 1 To TotalIFD
        IDTag = Buffer(Pos) + (Buffer(Pos + 1) * (2 ^ 8))
        
        Select Case IDTag
        
        Case 256 ' 100H
            ' Width
            ImgHeader.Width = Buffer(Pos + 8) + (Buffer(Pos + 9) * (2 ^ 8)) + (Buffer(Pos + 10) * (2 ^ 16)) + (Buffer(Pos + 11) * (2 ^ 24))
                
        Case 257 ' 101H
            ' Height
            ImgHeader.Height = Buffer(Pos + 8) + (Buffer(Pos + 9) * (2 ^ 8)) + (Buffer(Pos + 10) * (2 ^ 16)) + (Buffer(Pos + 11) * (2 ^ 24))
                
        Case 282 ' 11AH
            ' X Resolution
            Offset = Buffer(Pos + 8) + (Buffer(Pos + 9) * (2 ^ 8)) + (Buffer(Pos + 10) * (2 ^ 16)) + (Buffer(Pos + 11) * (2 ^ 24))
            ReDim BufferDPI(0 To 7)
            Get #FileNumber, Offset + 1, BufferDPI
            Resolution = (BufferDPI(0) + (BufferDPI(1) * (2 ^ 8)) + (BufferDPI(2) * (2 ^ 16)) + (BufferDPI(3) * (2 ^ 24))) / _
                        (BufferDPI(4) + (BufferDPI(5) * (2 ^ 8)) + (BufferDPI(6) * (2 ^ 16)) + (BufferDPI(7) * (2 ^ 24)))
                                
        Case 296 ' 128H
            ' Unit resolution 1=Centimeters and 2=Inchs
            ResolutionUnit = (Buffer(Pos + 8) + (Buffer(Pos + 9) * (2 ^ 8)) + (Buffer(Pos + 10) * (2 ^ 16)) + (Buffer(Pos + 11) * (2 ^ 24)))
        
        End Select
        
        Pos = Pos + 12
    Next F
     
    ' If is in centimeters convert to inchs
    If ResolutionUnit = 2 Then
        ImgHeader.DPI = Resolution
    Else
        ImgHeader.DPI = Int(Resolution * 2.54)
    End If

    ImgHeader.BitsPerPixel = 1
    GetHeaderTIF = True
    
End Function
Private Function GetHeaderJPG() As Boolean
    
    ' ******************************************************************
    ' * Until now I just test with JFIF e o EXIF formats.
    ' * The JPEG hearder work with markers each marker have 4 firts byth
    ' * with a type and the after 2 bytes is lenght marker, the size is
    ' * it self inclued. Each marker always begin with &HFF.
    ' *
    ' * More details visit - http://www.obrador.com/essentialjpeg/headerinfo.htm
    ' ******************************************************************
    
    On Error Resume Next
    
    ' Get JPEG type
    ReDim Buffer(0 To 3)
    Get #FileNumber, , Buffer
    
    ' Loop each marker header
    Do While Buffer(0) = &HFF
        Select Case Buffer(1)
        
        ' JFIF Marker
        Case &HE0
            ' Get the data
            ReDim Buffer(0 To (Buffer(2) * (2 ^ 8) + Buffer(3)) - 3)
            Get #FileNumber, , Buffer
       
            ' Resolution Unit
            If Buffer(7) = 1 Then
                ' Resolution in inch
                ImgHeader.DPI = Buffer(8) * (2 ^ 8) + Buffer(9)
            ElseIf Buffer(7) = 2 Then
                ' Resolution in centimeters - Convert to inch
                ImgHeader.DPI = (Buffer(8) * (2 ^ 8) + Buffer(9)) * 2.54
            End If
        
        Case &HC0
            ReDim Buffer(0 To (Buffer(2) * (2 ^ 8) + Buffer(3)) - 3)
            Get #FileNumber, , Buffer
    
            ' Height
            ImgHeader.Height = Buffer(1) * (2 ^ 8) + Buffer(2)
                
            ' Width
            ImgHeader.Width = Buffer(3) * (2 ^ 8) + Buffer(4)
            
            If Buffer(5) = 1 Then ImgHeader.IsGray = True
        
        Case Else
            ' Marker not used
            ReDim Buffer(0 To (Buffer(2) * (2 ^ 8) + Buffer(3)) - 3)
            Get #FileNumber, , Buffer
            
        End Select
        
        ReDim Buffer(0 To 3)
        Get #FileNumber, , Buffer
    Loop
    
    ' If no DPI assingn one
    If ImgHeader.DPI = 0 Then ImgHeader.DPI = 1
    
    ' In my case all jpeg must be 24 bits see the especification to other cases
    ImgHeader.BitsPerPixel = 24
    
    GetHeaderJPG = True
    
End Function
Property Get DPI() As Integer

    DPI = ImgHeader.DPI
    
End Property

Property Get IsGray() As Boolean

    IsGray = ImgHeader.IsGray
    
End Property

Property Get Width() As Long

    Width = ImgHeader.Width
    
End Property
Property Get Height() As Long

    Height = ImgHeader.Height
    
End Property
Property Get BitsPerPixel() As Byte

    BitsPerPixel = ImgHeader.BitsPerPixel
    
End Property
