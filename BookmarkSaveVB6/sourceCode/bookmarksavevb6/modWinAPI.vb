Imports System.Drawing
Imports System.Runtime.InteropServices
'---------------------------------------------------------------------
'
'WINAPI definitions for BookmarkSave project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------


Public Class WinAPI
    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function GetFocus() As IntPtr
        ' Leave function empty    
    End Function

    Private Declare Function GetClassNameInt Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As System.IntPtr, _
        ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer

    Private Declare Function GetWindowTextInt Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpString As String, ByVal cch As Integer) As Integer



    <DllImport("gdi32.dll", SetLastError:=True, CharSet:=CharSet.Auto, EntryPoint:="GetTextExtentPoint32A")> _
    Public Shared Function GetTextExtentPoint32(ByVal hDC As Integer, ByVal lpString As String, ByVal cbString As Integer, ByRef lpSize As apiSIZE) As Integer
    End Function


    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Shared Function GetDC(ByVal hWnd As Integer) As Integer
    End Function


    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Shared Function ReleaseDC(ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
    End Function


    <DllImport("gdi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Shared Function SelectObject(ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    End Function


    Public Shared Function GetClassName(hwnd As Integer) As String
        'some static vars to speed up things, this func will be called many times
        Dim sBuffer As String = Space(256)

        Dim r = GetClassNameInt(hwnd, sBuffer, 129&)
        If r > 0 Then
            Return Left(sBuffer, r)
        Else
            Return ""
        End If
    End Function


    Public Shared Function GetWindowText(hwnd As Integer, Optional Classname As Boolean = False) As String
        'some static vars to speed up things, this func will be called many times
        Dim sBuffer As String = Space(256)
        Dim r = GetWindowTextInt(hwnd, sBuffer, 129&)
        If r > 0 Then
            Return Left(sBuffer, r)
        Else
            Return ""
        End If
    End Function
End Class


Public Structure apiSIZE
    Public cx As Integer
    Public cy As Integer
End Structure


Public Class User32
    Public Const SPI_GETNONCLIENTMETRICS As Integer = 41
    Public Const LF_FACESIZE As Integer = 32


    <StructLayout(LayoutKind.Sequential)> _
    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure


    Declare Function GetWindowRect Lib "user32.dll" ( _
                ByVal hwnd As IntPtr, _
                ByRef lpRect As RECT) As Int32

    Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As IntPtr, ByRef pt As POINT) As Integer

    Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As IntPtr, ByRef pt As POINT) As Integer
End Class
