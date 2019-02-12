Imports System.Text
'---------------------------------------------------------------------
'
'Various Utility functions and API definitions for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


Module modUtility
    Public Const GW_HWNDNEXT As Integer = 2
    Public Const GW_CHILD As Integer = 5
    Public Const CB_GETCOUNT As Integer = &H146
    Public Const CB_GETLBTEXT As Integer = &H148
    Public Const CB_RESETCONTENT As Integer = &H14B

    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
       (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

    Public Declare Function GetDesktopWindow Lib "user32" () As Integer

    Public Declare Function GetWindow Lib "user32" _
       (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer

    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
       (ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer

    Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
       (ByVal hwnd As Integer, ByVal lpClassName As String, _
        ByVal nMaxCount As Integer) As Integer

    Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, _
        lParam As Integer) As Integer

    Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, _
        lParam As String) As Integer

    Public Declare Function GetParent Lib "user32" (ByVal hwnd As Integer) As Integer

    'New for Keyboard Hooks
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Integer, lpdwProcessId As Integer) As Integer


    Public Function Deserialize(Of t)(ByVal XMLString As String) As t
        Dim stream = New System.IO.StringReader(XMLString)
        Dim reader = System.Xml.XmlReader.Create(stream)

        Dim serializer = New System.Runtime.Serialization.DataContractSerializer(GetType(t))
        Dim result = DirectCast(serializer.ReadObject(reader), t)

        Return result
    End Function


    Public Function Serialize(Of t)(ByVal Obj As t) As String
        Dim stream = New System.IO.MemoryStream
        Dim xmlsettings = New Xml.XmlWriterSettings
        xmlsettings.Indent = True
        xmlsettings.Encoding = New UTF8Encoding
        Dim writer = System.Xml.XmlWriter.Create(stream, xmlsettings)

        Dim serializer = New System.Runtime.Serialization.DataContractSerializer(GetType(t))
        serializer.WriteObject(writer, Obj)
        writer.Flush()

        Return xmlsettings.Encoding.GetString(stream.ToArray())
    End Function


    Public Function GetComboRefFromFindWindow(ByVal hWndStart As Integer) As Integer

        Dim hwnd As Integer
        Dim sClassname As String
        Dim r As Integer
        Dim lBlnFoundIt As Boolean
        Dim lStrData As String
        Dim lLngItemCount As Integer
        Dim sWindowText As String
        'need to start the window parent as VB

        'Get first child window
        hwnd = GetWindow(hWndStart, GW_CHILD)

        'Search children by recursion
        Do Until hwnd = 0

            sClassname = Space(255)
            r = GetClassName(hwnd, sClassname, 255)
            sClassname = Left(sClassname, r)

            'Get the window text and class name
            sWindowText = Space(255)
            r = GetWindowText(hwnd, sWindowText, 255)
            sWindowText = Left(sWindowText, r)

            If sClassname = "ComboBox" Then
                lBlnFoundIt = True
                'get the handle to the edit portion  'of the combo control
                'check value could be wrong combo
                lLngItemCount = SendMessageLong(hwnd, CB_GETCOUNT, 0&, 0&)
                If lLngItemCount = 3 Then
                    lStrData = ""
                    lStrData = Space(255)
                    r = SendMessageStr(hwnd, CB_GETLBTEXT, 0&, lStrData)
                    If mfStrStripSpaces(lStrData) = "All" Then
                        lStrData = ""
                        lStrData = Space(255)
                        r = SendMessageStr(hwnd, CB_GETLBTEXT, 1&, lStrData)
                        If mfStrStripSpaces(lStrData) = "Down" Then
                            lStrData = ""
                            lStrData = Space(255)
                            r = SendMessageStr(hwnd, CB_GETLBTEXT, 2&, lStrData)
                            If mfStrStripSpaces(lStrData) = "Up" Then
                                'wrong one
                                lBlnFoundIt = False
                            End If
                        End If
                    End If

                    If lBlnFoundIt = True Then
                        Return hwnd
                    End If
                Else
                    'It the one we want
                    Return hwnd
                End If
            End If

            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
            'DoEvents()
        Loop
        Return 0
    End Function


    Private Function mfStrStripSpaces(ByVal pvStrIn As String) As String
        Dim lPos As Integer

        lPos = InStr(pvStrIn, Chr(0))
        If lPos Then
            Return Left(pvStrIn, lPos - 1)
        End If
        Return ""
    End Function


    Function FindWindowLike(ByVal hWndStart As Integer, _
                            WindowText As String, _
                            Classname As String, _
                            Optional ByVal pvBln_StartOver As Boolean = False) As Integer()

        'Hold the level of recursion and
        'hold the number of matching windows
        Static level As Integer

        Dim hwnds() As Integer = {}
        Try

            Dim hwnd As Integer
            Dim sWindowText As String
            Dim sClassname As String
            Dim r As Integer

            'Initialize if necessary
            If level = 0 Then
                If hWndStart = 0 Then hWndStart = GetDesktopWindow()
            End If

            'Increase recursion counter
            level = level + 1

            'Get first child window
            hwnd = GetWindow(hWndStart, GW_CHILD)

            Do Until hwnd = 0
                'Search children by recursion
                Dim hw = FindWindowLike(hwnd, WindowText, Classname)
                If hw.Count > 0 Then hwnds = hwnds.Concat(hw).ToArray

                'Get the window text and class name
                sWindowText = Space(256)
                r = GetWindowText(hwnd, sWindowText, 255)
                If r > 0 Then
                    sWindowText = Left(sWindowText, r)
                Else
                    sWindowText = vbNullString
                End If
                sClassname = Space(255)
                r = GetClassName(hwnd, sClassname, 255)
                If r > 0 Then
                    sClassname = Left(sClassname, r)
                Else
                    sClassname = vbNullString
                End If

                'Check that window matches the search parameters
                If (sWindowText Like WindowText) And (sClassname Like Classname) Then
                    hwnds = hwnds.Concat({hwnd}).ToArray
                End If
                'Get next child window
                hwnd = GetWindow(hwnd, GW_HWNDNEXT)
            Loop
        Catch
        End Try

        'Decrement recursion counter
        level = level - 1

        'Return the number of windows found
        Return hwnds
    End Function

End Module
