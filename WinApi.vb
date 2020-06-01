Imports System.Runtime.InteropServices

Public Class WinApi

    ' Everything needed to allow us to call the relevant Win32 APIs from .NET code.
    ' See links to native API definitions on each item for further details

    Public Const SERVICE_WIN32_OWN_PROCESS As UInteger = &H10
    Public Const SERVICE_ERROR_NORMAL As UInteger = 1
    Public Const SERVICE_AUTO_START As UInteger = 2
    Public Const SERVICE_DEMAND_START As UInteger = 3
    Public Const SERVICE_DISABLED As UInteger = 4

    'https://docs.microsoft.com/en-us/windows/win32/services/service-security-and-access-rights#access-rights-for-a-service
    <Flags()> _
    Public Enum SERVICE_RIGHTS As UInteger
        SERVICE_QUERY_CONFIG = 1
        SERVICE_CHANGE_CONFIG = 2
        SERVICE_QUERY_STATUS = 4
        SERVICE_ENUMERATE_DEPENDENTS = 8
        SERVICE_START = &H10
        SERVICE_STOP = &H20
        SERVICE_INTERROGATE = &H80
        SERVICE_USER_DEFINED_CONTROL = &H100
        SERVICE_PAUSE_CONTINUE = &H40
        SERVICE_ALL_ACCESS = &HF01FF
        ACCESS_SYSTEM_SECURITY = &H1000000
        DELETE = &H10000
        WRITE_DAC = &H40000
        WRITE_OWNER = &H80000
        SYNCHRONIZE = &H100000
    End Enum

    'https://docs.microsoft.com/en-us/windows/win32/services/service-security-and-access-rights
    <Flags()> _
    Public Enum SCM_RIGHTS As UInteger
        SC_MANAGER_CONNECT = 1
        SC_MANAGER_CREATE_SERVICE = 2
        SC_MANAGER_ENUMERATE_SERVICE = 4
        SC_MANAGER_LOCK = 8
        SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
        SC_MANAGER_QUERY_LOCK_STATUS = &H10
        SC_MANAGER_ALL_ACCESS = &HF003F
        DELETE = &H10000
        WRITE_DAC = &H40000
        WRITE_OWNER = &H80000
        SYNCHRONIZE = &H100000
    End Enum

    'https://docs.microsoft.com/en-us/windows/win32/api/winsvc/nf-winsvc-startservicew
    <DllImport("advapi32.dll", EntryPoint:="StartServiceW", SetLastError:=True)> _
    Public Shared Function StartService(ByVal hService As IntPtr, ByVal dwNumServiceArgs As UInteger, _
                                        ByVal lpServiceArgVectors As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/winsvc/nf-winsvc-createservicew
    <DllImport("advapi32.dll", EntryPoint:="CreateServiceW", SetLastError:=True)> _
    Public Shared Function CreateService(ByVal hSCManager As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpServiceName As String, _
                                         <MarshalAs(UnmanagedType.LPWStr)> ByVal lpDisplayName As String, ByVal dwDesiredAccess As UInteger, ByVal dwServiceType As UInteger, ByVal dwStartType As UInteger,
                                         ByVal dwErrorControl As UInteger, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpBinaryPathName As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpLoadOrderGroup As String, _
                                         ByVal lpdwTagId As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpDependencies As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpServiceStartName As String, _
                                         <MarshalAs(UnmanagedType.LPWStr)> ByVal lpPassword As String) As IntPtr
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/winsvc/nf-winsvc-openscmanagerw
    <DllImport("advapi32.dll", EntryPoint:="OpenSCManagerW", SetLastError:=True)> _
    Public Shared Function OpenSCManager(<MarshalAs(UnmanagedType.LPWStr)> ByVal lpMachineName As String, _
                                         <MarshalAs(UnmanagedType.LPWStr)> ByVal lpDatabaseName As String, _
                                         ByVal dwDesiredAccess As UInteger) As IntPtr
    End Function

    'https://docs.microsoft.com/en-us/windows/win32/api/winsvc/nf-winsvc-closeservicehandle
    <DllImport("advapi32.dll", EntryPoint:="CloseServiceHandle", SetLastError:=True)> _
    Public Shared Function CloseServiceHandle(ByVal hSCObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", EntryPoint:="OpenServiceW", SetLastError:=True)> _
    Public Shared Function OpenService(ByVal hSCManager As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpServiceName As String, _
                                        ByVal dwDesiredAccess As UInteger) As IntPtr
    End Function

End Class
