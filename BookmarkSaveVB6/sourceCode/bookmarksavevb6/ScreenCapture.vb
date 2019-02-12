Imports System
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging


Namespace ScreenShot
    ''' <summary>
    ''' Simple functions for screen captures in VB
    ''' From a post here
    ''' http://www.vbforums.com/showthread.php?t=385497
    ''' Adapted and extended slightly by Darin Higgins Oct, 2011
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Capture
        ''' <summary>
        ''' Captures the primary monitor screen
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Screen() As Image
            Return Window(User32.GetDesktopWindow())
        End Function


        ''' <summary>
        ''' Captures the entire desktop when multiple monitors are present
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EntireDesktop() As Image
            Dim base As IntPtr = User32.FindWindow("Progman", "Program Manager")
            Dim b2 As IntPtr = User32.FindWindowEx(base, vbNullString, "SHELLDLL_DefView", vbNullString)
            b2 = User32.FindWindowEx(b2, vbNullString, "SysListView32", vbNullString)

            Return Window(b2)
        End Function


        ''' <summary>
        ''' Return image of a specific window
        ''' </summary>
        ''' <param name="hwnd"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Window(ByVal hwnd As IntPtr) As Image
            Dim wRect As New User32.RECT
            User32.GetWindowRect(hwnd, wRect)
            Dim Rect = New Rectangle(0, 0, wRect.right - wRect.left, wRect.bottom - wRect.top)

            Return windowRect(hwnd, Rect)
        End Function


        ''' <summary>
        ''' Retrieves a rectangle (in window's client coords) from the target window
        ''' </summary>
        ''' <param name="hwnd"></param>
        ''' <param name="Rect"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ClientWindowRect(ByVal hwnd As IntPtr, ByVal Rect As Rectangle) As Image
            '---- the rectangle is in client coords
            '     so add in the NC left and top to adjust it to window coords
            '---- get 0,0 client coords in screen coords
            Dim pt As User32.POINT
            pt.X = 0 : pt.Y = 0
            User32.ClientToScreen(hwnd, pt)

            '---- get the window rect (including NC area) in screen coords
            Dim wRect As New User32.RECT
            User32.GetWindowRect(hwnd, wRect)

            '---- to convert the Rect from Client Coords into Window Coords
            Rect.Offset(pt.X - wRect.left, pt.Y - wRect.top)

            '---- and capture based on the WINDOW coords
            Return WindowRect(hwnd, Rect)
        End Function


        ''' <summary>
        ''' Return shot of a portion of a specific window
        ''' </summary>
        ''' <param name="hwnd"></param>
        ''' <param name="Rect"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function WindowRect(ByVal hwnd As IntPtr, ByVal Rect As Rectangle) As Image
            Dim SRCCOPY As Integer = &HCC0020
            ' get te hDC of the target window
            Dim hdcSrc As IntPtr = User32.GetWindowDC(hwnd)
            ' get the size
            ' create a device context we can copy to
            Dim hdcDest As IntPtr = GDI32.CreateCompatibleDC(hdcSrc)
            ' create a bitmap we can copy it to,
            ' using GetDeviceCaps to get the width/height
            Dim hBitmap As IntPtr = GDI32.CreateCompatibleBitmap(hdcSrc, Rect.Width, Rect.Height)
            ' select the bitmap object
            Dim hOld As IntPtr = GDI32.SelectObject(hdcDest, hBitmap)
            ' bitblt over
            GDI32.BitBlt(hdcDest, 0, 0, Rect.Width, Rect.Height, hdcSrc, Rect.Left, Rect.Top, SRCCOPY)
            ' restore selection
            GDI32.SelectObject(hdcDest, hOld)
            ' clean up 
            GDI32.DeleteDC(hdcDest)
            User32.ReleaseDC(hwnd, hdcSrc)

            ' get a .NET image object for it
            Dim img As Image = Image.FromHbitmap(hBitmap)
            ' free up the Bitmap object
            GDI32.DeleteObject(hBitmap)

            Return img
        End Function


        ''' <summary>
        ''' Capture shot of a window and save to a file
        ''' </summary>
        ''' <param name="hwnd"></param>
        ''' <param name="filename"></param>
        ''' <param name="format"></param>
        ''' <remarks></remarks>
        Public Shared Sub WindowToFile(ByVal hwnd As IntPtr, ByVal filename As String, ByVal format As ImageFormat)
            Dim img As Image = Window(hwnd)
            img.Save(filename, format)
        End Sub


        ''' <summary>
        ''' Capture shot of primary desktop monitor and save to file
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="format"></param>
        ''' <remarks></remarks>
        Public Shared Sub ScreenToFile(ByVal filename As String, ByVal format As ImageFormat)
            Dim img As Image = Screen()
            img.Save(filename, format)
        End Sub


        ''' <summary>
        ''' Capture a rectangular portion of the desktop
        ''' </summary>
        ''' <param name="CapRect"></param>
        ''' <param name="CapRectWidth"></param>
        ''' <param name="CapRectHeight"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DeskTopRect(ByVal CapRect As Rectangle, ByVal CapRectWidth As Integer, ByVal CapRectHeight As Integer) As Bitmap
            '/ Returns BitMap of the region of the desktop, similar to CaptureWindow, but can be used to 
            '/ create a snapshot of the desktop when no handle is present, by passing in a rectangle 
            '/ Grabs snapshot of entire desktop, then crops it using the passed in rectangle's coordinates
            Dim bmpImage As New Bitmap(Screen())
            Dim bmpCrop As New Bitmap(CapRectWidth, CapRectHeight, bmpImage.PixelFormat)
            Dim recCrop As New Rectangle(CapRect.X, CapRect.Y, CapRectWidth, CapRectHeight)
            Dim gphCrop As Graphics = Graphics.FromImage(bmpCrop)
            Dim recDest As New Rectangle(0, 0, CapRectWidth, CapRectHeight)
            gphCrop.DrawImage(bmpImage, recDest, recCrop.X, recCrop.Y, recCrop.Width, _
              recCrop.Height, GraphicsUnit.Pixel)
            Return bmpCrop
        End Function


        ''' <summary>
        ''' Helper GDI p/invoke functions
        ''' </summary>
        ''' <remarks></remarks>
        Private Class GDI32
            Public SRCCOPY As Integer = &HCC0020
            ' BitBlt dwRop parameter
            Declare Function BitBlt Lib "gdi32.dll" ( _
                ByVal hDestDC As IntPtr, _
                ByVal x As Int32, _
                ByVal y As Int32, _
                ByVal nWidth As Int32, _
                ByVal nHeight As Int32, _
                ByVal hSrcDC As IntPtr, _
                ByVal xSrc As Int32, _
                ByVal ySrc As Int32, _
                ByVal dwRop As Int32) As Int32

            Declare Function CreateCompatibleBitmap Lib "gdi32.dll" ( _
                ByVal hdc As IntPtr, _
                ByVal nWidth As Int32, _
                ByVal nHeight As Int32) As IntPtr

            Declare Function CreateCompatibleDC Lib "gdi32.dll" ( _
                ByVal hdc As IntPtr) As IntPtr

            Declare Function DeleteDC Lib "gdi32.dll" ( _
                ByVal hdc As IntPtr) As Int32

            Declare Function DeleteObject Lib "gdi32.dll" ( _
                ByVal hObject As IntPtr) As Int32

            Declare Function SelectObject Lib "gdi32.dll" ( _
                ByVal hdc As IntPtr, _
                ByVal hObject As IntPtr) As IntPtr
        End Class 'GDI32


        ''' <summary>
        ''' Helper User32 p/Invoke functions
        ''' </summary>
        ''' <remarks></remarks>
        Private Class User32
            <StructLayout(LayoutKind.Sequential)> _
            Public Structure RECT
                Public left As Integer
                Public top As Integer
                Public right As Integer
                Public bottom As Integer
            End Structure 'RECT

            <StructLayout(LayoutKind.Sequential)> _
            Public Structure POINT
                Public X As Integer
                Public Y As Integer
            End Structure

            Declare Function GetDesktopWindow Lib "user32.dll" () As IntPtr

            Declare Function GetWindowDC Lib "user32.dll" ( _
                ByVal hwnd As IntPtr) As IntPtr

            Declare Function ReleaseDC Lib "user32.dll" ( _
                ByVal hwnd As IntPtr, _
                ByVal hdc As IntPtr) As Int32

            Declare Function GetWindowRect Lib "user32.dll" ( _
                ByVal hwnd As IntPtr, _
                ByRef lpRect As RECT) As Int32

            Declare Function FindWindowEx Lib "User32.dll" ( _
                ByVal hwndParent As IntPtr, _
                ByVal hwndChild As IntPtr, _
                ByVal className As String, _
                ByVal caption As String) As IntPtr

            Declare Function FindWindow Lib "User32.dll" ( _
                ByVal className As String, _
                ByVal caption As String) As IntPtr

            Declare Function ScreenToClient Lib "user32.dll" ( _
                ByVal hWnd As IntPtr, ByRef pt As POINT) As Integer

            Declare Function ClientToScreen Lib "user32.dll" ( _
                ByVal hWnd As IntPtr, ByRef pt As POINT) As Integer

        End Class
    End Class
End Namespace
