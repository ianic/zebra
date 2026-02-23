' Visual Basic 3.0 declares file for
'
'     azip16.dll    addZIP 16-bit compression library
'     aunzip16.dll  addUNZIP 16-bit decompression library
'
' Copyright © 1996, Stephen Darlington. All rights reserved.
' Written by Stephen Darlington                  May 1996
                                      
' compression functions
Declare Function addZIP Lib "azip16.dll" () As Integer
Declare Function addZIP_ArchiveName Lib "azip16.dll" (ByVal lpFile As String) As Integer
Declare Function addZIP_Encrypt Lib "azip16.dll" (ByVal lpPassword As String) As Integer
Declare Function addZIP_Exclude Lib "azip16.dll" (ByVal lpFiles As String) As Integer
Declare Function addZIP_ExcludeListFile Lib "azip16.dll" (ByVal lpFile As String) As Integer
Declare Function addZIP_Include Lib "azip16.dll" (ByVal lpFiles As String) As Integer
Declare Function addZIP_IncludeListFile Lib "azip16.dll" (ByVal lpFile As String) As Integer
Declare Function addZIP_Overwrite Lib "azip16.dll" (ByVal iFlag As Integer) As Integer
Declare Function addZIP_Recurse Lib "azip16.dll" (ByVal iFlag As Integer) As Integer

Declare Function addZIP_SaveStructure Lib "azip16.dll" (ByVal iFlag As Integer) As Integer
Declare Function addZIP_SetCompressionLevel Lib "azip16.dll" (ByVal iFlag As Integer) As Integer
Declare Function addZIP_SetWindowHandle Lib "azip16.dll" (ByVal hwnd As Integer) As Integer
Declare Function addZIP_SetParentWindowHandle Lib "azip16.dll" (ByVal hwnd As Integer) As Integer
Declare Function addZIP_View Lib "azip16.dll" (ByVal bFlag As Integer) As Integer

' decompression functions
Declare Function addUNZIP Lib "aunzip16.dll" () As Integer
Declare Function addUNZIP_ArchiveName Lib "aunzip16.dll" (ByVal lpFile As String) As Integer
Declare Function addUNZIP_Decrypt Lib "aunzip16.dll" (ByVal lpPassword As String) As Integer
Declare Function addUNZIP_Exclude Lib "aunzip16.dll" (ByVal lpFiles As String) As Integer
Declare Function addUNZIP_ExcludeListFile Lib "aunzip16.dll" (ByVal lpFile As String) As Integer
Declare Function addUNZIP_ExtractTo Lib "aunzip16.dll" (ByVal lpPath As String) As Integer
Declare Function addUNZIP_Freshen Lib "aunzip16.dll" (ByVal bFlag As Integer) As Integer
Declare Function addUNZIP_Include Lib "aunzip16.dll" (ByVal lpFiles As String) As Integer
Declare Function addUNZIP_IncludeListFile Lib "aunzip16.dll" (ByVal lpFile As String) As Integer
Declare Function addUNZIP_Overwrite Lib "aunzip16.dll" (ByVal iFlag As Integer) As Integer
'Declare Function addUNZIP_Register Lib "aunzip16.dll" (??????) As Integer
Declare Function addUNZIP_RestoreStructure Lib "aunzip16.dll" (ByVal bFlag As Integer) As Integer
Declare Function addUNZIP_SetParentWindowHandle Lib "aunzip16.dll" (ByVal hwnd As Integer) As Integer
Declare Function addUNZIP_SetWindowHandle Lib "aunzip16.dll" (ByVal hwnd As Integer) As Integer
Declare Function addUNZIP_Test Lib "aunzip16.dll" (ByVal bFlag As Integer) As Integer
Declare Function addUNZIP_Update Lib "aunzip16.dll" (ByVal bFlag As Integer) As Integer
Declare Function addUNZIP_View Lib "aunzip16.dll" (ByVal bFlag As Integer) As Integer

' global constants for addZIP_Overwrite and addUNZIP_Overwrite
Global Const OVERWRITE_QUERY = 10
Global Const OVERWRITE_ALL = 11
Global Const OVERWRITE_NONE = 12

' global constants for addZIP_SetCompressionLevel
Global Const COMPRESSION_NONE = 0
Global Const COMPRESSION_MINIMUM = 1
Global Const COMPRESSION_NORMAL = 2
Global Const COMPRESSION_MAXIMUM = 3

' global constants for addZIP_SaveStructure
Global Const SAVE_FILENAME_ONLY = 0
Global Const SAVE_RELATIVE_PATH = 1
Global Const SAVE_ABSOLUTE_PATH = 2

Function GetFileCompressedSize (cFrom As String) As Long
    GetFileCompressedSize = Val(GetPiece(cFrom, "|", 7))
End Function

Function GetFileCompressionRatio (cFrom As String) As Integer
    GetFileCompressionRatio = Val(GetPiece(cFrom, "|", 8))
End Function

Function GetFileName (cFrom As String) As String
    GetFileName = GetPiece(cFrom, "|", 5)
End Function

Function GetFileOriginalSize (cFrom As String) As Long
    GetFileOriginalSize = Val(GetPiece(cFrom, "|", 6))
    Debug.Print Val(GetPiece(cFrom, "|", 6))
End Function

Function GetFilePath (cFrom As String) As String
    GetFilePath = GetPiece(cFrom, "|", 4)
End Function

Function GetPiece (from As String, delim As String, Index As Integer) As String
    Dim temp$
    Dim Count As Integer
    Dim Where As Integer
    '
    temp$ = from & delim
    Where = InStr(temp$, delim)
    Count = 0
    Do While (Where > 0)
        Count = Count + 1
        If (Count = Index) Then
            GetPiece = Left$(temp$, Where - 1)
            Exit Function
        End If
        temp$ = Right$(temp$, Len(temp$) - Where)
        Where = InStr(temp$, delim)
    Loop
    If (Count = 0) Then
        GetPiece = from
    Else
        GetPiece = ""
    End If
End Function

