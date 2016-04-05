Attribute VB_Name = "MyExports"

Declare Function lenCString Lib "Kernel32" Alias "lstrlenA" (lpString As Long) As Long
Declare Function CopyCString Lib "Kernel32" Alias "lstrcpynA" (ByVal lpStringDestination As String, lpStringSource As Long, ByVal lngMaxLength As Long) As Long

Public Sub Main()
End Sub

Public Function NumberString(ByVal lngAnyNumber As Long, ByRef lngStringPtr As Long) As String
    NumberString = lngAnyNumber & CStringToVBString(lngStringPtr)
End Function

Function CStringToVBString(lpCString As Long) As String
    Dim lenString As Long, sBuffer As String, lpBuffer As Long, lngStringPointer As Long, refStringPointer As Long
    If lpCString = 0 Then
        CStringToVBString = vbNullString
    Else
        lenString = lenCString(lpCString)
        sBuffer = String$(lenString + 1, 0) 'buffer has one extra byte for terminator
        lpBuffer = CopyCString(sBuffer, lpCString, lenString + 1)
        CStringToVBString = sBuffer
    End If
End Function

