Attribute VB_Name = "UseDLL"

Declare Function NumberString Lib "NormalDLL.DLL" (ByVal lngNumber As Long, ByVal sString As String) As Long

Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) 'thank you bruce

Sub Main()
    Dim lngAnyNumber As Long, sAnyString As String, lngResultPtr As Long, sResult As String
    lngAnyNumber = Val(InputBox("Enter any number:", "Normal DLL arg 1"))
    If lngAnyNumber = 0 Then Exit Sub
    sAnyString = InputBox("Enter any string:", "Normal DLL arg 2")
    If sAnyString = "" Then Exit Sub
    lngResultPtr = NumberString(lngAnyNumber, sAnyString)
    CopyMemory ByVal VarPtr(sResult), lngResultPtr, 4
    MsgBox "Result string is: " & sResult
End Sub

'The only tricky move here is the copy memory statement. There are various ways to get
'the BSTR from the DLL into the BSTR called "sResult" here. I have chosen to do it by
'copying a pointer returned by the DLL call into the sResult.

'Note that some kind of pointer connection has to be done, because VB normally assumes
'that a string coming from a DLL call is a C-string and tries to translate it. You could
'potentially avoid this conversion problem with a type library--but then you would have
'to write the type library. If you are making a big "normal" DLL for VB you should take
'this course and learn how to write a type library for it. If you do that VB will just
'suck in BSTR and other OLE types like variants naturally and you can avoid pointer
'manipulation.
