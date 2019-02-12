Imports Microsoft.Win32
'---------------------------------------------------------------------
'
'DLL Registration functions for Bookmarksave Project
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

'---- Write additional info the registry when the user uses REGSVR32 on this dll
'     to actually register it with VB6
'     NOTE. I let failures fall out because REgsvr32 will report them
Partial Public Class DLLRegisterFunctions

    ''' <summary>
    ''' Register info for VB6
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AdditionalRegistration()
        Using regKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Visual Basic\6.0\Addins", True)
            Using subKey = regKey.CreateSubKey("BookmarkSave.Connect")
                subKey.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord)
                subKey.SetValue("Description", "BookmarkSave VB6 addin")
                subKey.SetValue("FriendlyName", "BookmarkSave VB6 addin")
                subKey.SetValue("LoadBehavior", 7, RegistryValueKind.DWord)
            End Using
        End Using
    End Sub


    ''' <summary>
    ''' Additional unregistration stuff for a VB6 addin
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AdditionalUnregistration()
        Using regKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Visual Basic\6.0\Addins", True)
            regKey.DeleteSubKeyTree("BookmarkSave.Connect")
        End Using
    End Sub
End Class
