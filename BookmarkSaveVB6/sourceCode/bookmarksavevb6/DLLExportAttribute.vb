Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

'=============================================================================
'=============================================================================
'This DLL exists only to provide easy access to the DLLExportAttribute
'without having to include the source file in your application
'Your application needs to either include the source for the DLLExportAttribute
'or include a reference to this dll to use it.
'
'When DLLExport is run on the target DLL, however, references to this 
'dll will automatically be removed, so there is no need to include this
'dll in your shipping fileset.
'=============================================================================
'=============================================================================


Namespace DllExport
    ''' <summary>
    ''' Attribute added to a static method to export it
    ''' </summary>
    <AttributeUsage(AttributeTargets.Method)> _
    Public Class DllExportAttribute
        Inherits Attribute
        
        ''' <summary>
        ''' Constructor 1
        ''' </summary>
        ''' <param name="exportName"></param>
        Public Sub New(exportName As String)
            Me.New(exportName, System.Runtime.InteropServices.CallingConvention.StdCall)
        End Sub


        ''' <summary>
        ''' Constructor 2
        ''' </summary>
        ''' <param name="exportName"></param>
        ''' <param name="callingConvention"></param>
        Public Sub New(exportName As String, callingConvention As CallingConvention)
            _ExportName = exportName
            _CallingConvention = callingConvention
        End Sub
        Private _ExportName As String


        ''' <summary>
        ''' Get the export name, or null to use the method name
        ''' </summary>
        Public ReadOnly Property ExportName() As String
            Get
                Return _ExportName
            End Get
        End Property


        ''' <summary>
        ''' Get the calling convention
        ''' </summary>
        Public ReadOnly Property CallingConvention() As String
            Get
                Select Case _CallingConvention
                    Case System.Runtime.InteropServices.CallingConvention.Cdecl
                        Return GetType(CallConvCdecl).FullName

                    Case System.Runtime.InteropServices.CallingConvention.FastCall
                        Return GetType(CallConvFastcall).FullName

                    Case System.Runtime.InteropServices.CallingConvention.StdCall
                        Return GetType(CallConvStdcall).FullName

                    Case System.Runtime.InteropServices.CallingConvention.ThisCall
                        Return GetType(CallConvThiscall).FullName

                    Case System.Runtime.InteropServices.CallingConvention.Winapi
                        Return GetType(CallConvStdcall).FullName
                    Case Else

                        Return ""
                End Select
            End Get
        End Property
        Private _CallingConvention As CallingConvention

    End Class
End Namespace
