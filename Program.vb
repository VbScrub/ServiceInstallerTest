Imports System.Runtime.InteropServices
Imports System.ComponentModel

Module Program

    Sub Main()

        If My.Application.CommandLineArgs.Count < 2 Then
            Console.WriteLine("Usage:" & Environment.NewLine &
                              "ServiceInstallTest.exe ServiceName ExePath")
            Exit Sub
        End If
        ' Get command line args
        Dim ServiceName As String = My.Application.CommandLineArgs(0)
        Dim ExePath As String = My.Application.CommandLineArgs(1)
        Dim ScmHandle As IntPtr = IntPtr.Zero
        Dim ServiceHandle As IntPtr = IntPtr.Zero

        Try
            ' Connect to SCM (Service Control Manager) and get a handle that allows us to create services by specifying SC_MANAGER_CREATE_SERVICE. 
            ' Would fail with access denied here if we don't have permission to create services)
            ScmHandle = WinApi.OpenSCManager(Nothing, Nothing, WinApi.SCM_RIGHTS.SC_MANAGER_CREATE_SERVICE)
            If ScmHandle = IntPtr.Zero Then
                Throw New Win32Exception()
            End If

            Try
                ' Create a new service and get back a handle to the service (that can then be used with StartService function and other service functions). 
                ' We specify SERVICE_ALL_ACCESS so that this handle allows us full control over the service (to start it, stop it, delete it, etc) and for some
                ' reason this works. If we try to request this level of access afterwards by using the OpenService function, it fails with access denied
                ServiceHandle = WinApi.CreateService(ScmHandle,
                                                    ServiceName,
                                                    Nothing,
                                                    WinApi.SERVICE_RIGHTS.SERVICE_ALL_ACCESS,
                                                    WinApi.SERVICE_WIN32_OWN_PROCESS,
                                                    WinApi.SERVICE_DEMAND_START,
                                                    WinApi.SERVICE_ERROR_NORMAL,
                                                    ExePath,
                                                    Nothing, Nothing, Nothing, "NT Authority\System", Nothing)

                ' If we successfully created the service, attempt to start it using the handle we got back from CreateService. 
                If Not ServiceHandle = IntPtr.Zero Then
                    If WinApi.StartService(ServiceHandle, 0, IntPtr.Zero) Then
                        Console.Write("Service created and started successfully")
                    Else
                        Throw New Exception("Service has been created but was unable to start due to error: " & New System.ComponentModel.Win32Exception().Message)
                    End If
                Else
                    Throw New Exception("Unable to create service due to error: " & New System.ComponentModel.Win32Exception().Message)
                End If

            Finally
                If Not ServiceHandle = IntPtr.Zero Then
                    WinApi.CloseServiceHandle(ServiceHandle)
                End If
            End Try
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            If Not ScmHandle = IntPtr.Zero Then
                WinApi.CloseServiceHandle(ScmHandle)
            End If
        End Try
    End Sub


End Module
