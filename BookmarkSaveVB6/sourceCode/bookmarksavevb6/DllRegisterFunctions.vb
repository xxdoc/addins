Imports System.Runtime.InteropServices
Imports System.Reflection
'---------------------------------------------------------------------
'
'DLLRegisterServer and UnregisterServer entrypoints
'
'(c) 2012 Darin Higgins
'
'---------------------------------------------------------------------

''' <summary>
''' Provides Self-Registration functions for COM exposed .net assemblies
''' If you need additional Registration or Unregistration code (say
''' to write additional entries into the Registry), create a Partial Class
''' of the DLLRegisterFunctions class and define two parameterless subs
''' called AdditionalRegistration and AdditionalUnregistration
''' </summary>
''' <remarks></remarks>
Partial Public Class DllRegisterFunctions
    Public Const S_OK As Integer = 0
    Public Const SELFREG_E_TYPELIB = &H80040200
    Public Const SELFREG_E_CLASS = &H80040201


    <DllExport.DllExport("DllRegisterServer", CallingConvention.Cdecl)> _
    Public Shared Function DllRegisterServer() As Integer
        Try
            Dim asm = Assembly.LoadFile(Assembly.GetExecutingAssembly.Location)
            Dim regAsm = New RegistrationServices()
            Dim bResult = regAsm.RegisterAssembly(asm, AssemblyRegistrationFlags.SetCodeBase)

            Try
                '---- attempt to call out to any addition registration function
                Dim Reg = New DLLRegisterFunctions
                CallByName(Reg, "AdditionalRegistration", CallType.Method)
            Catch
                '---- the function couldn't be found latebound so just ignore the error
            End Try
            Return S_OK
        Catch ex As Exception
            Return SELFREG_E_TYPELIB
        End Try
    End Function


    <DllExport.DllExport("DllUnregisterServer", CallingConvention.Cdecl)> _
    Public Shared Function DllUnregisterServer() As Integer
        Try
            Dim asm = Assembly.LoadFile(Assembly.GetExecutingAssembly.Location)
            Dim regAsm = New RegistrationServices()
            Dim bResult = regAsm.UnregisterAssembly(asm)
            Try
                '---- attempt to call out to any addition unregistration function
                Dim reg = New DLLRegisterFunctions
                CallByName(reg, "AdditionalUnRegistration", CallType.Method)
            Catch
                '---- the function couldn't be found latebound so just ignore the error
            End Try
            Return S_OK
        Catch ex As Exception
            Return SELFREG_E_TYPELIB
        End Try
    End Function
End Class
