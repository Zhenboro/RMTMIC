Imports Microsoft.Win32
Module GlobalUses
    Public parameters As String
    Public DIRCommons As String = "C:\Users\" & Environment.UserName & "\AppData\Local\Zhenboro"
    Public DIRHome As String = DIRCommons & "\" & My.Application.Info.AssemblyName
End Module
Module Utility
    Public tlmContent As String
    Function AddToLog(ByVal from As String, ByVal content As String, Optional ByVal flag As Boolean = False) As String
        Try
            Dim OverWrite As Boolean = False
            If My.Computer.FileSystem.FileExists(DIRHome & "\" & My.Application.Info.AssemblyName & ".log") Then
                OverWrite = True
            End If
            Dim finalContent As String = Nothing
            If flag = True Then
                finalContent = " [!!!]"
            End If
            Dim Message As String = DateTime.Now.ToString("hh:mm:ss tt dd/MM/yyyy") & finalContent & " [" & from & "] " & content
            tlmContent = tlmContent & Message & vbCrLf
            Console.WriteLine("[" & from & "]" & finalContent & " " & content)
            Try
                My.Computer.FileSystem.WriteAllText(DIRHome & "\" & My.Application.Info.AssemblyName & ".log", vbCrLf & Message, OverWrite)
            Catch
            End Try
            Return finalContent & "[" & from & "]" & content
        Catch ex As Exception
            Console.WriteLine("[AddToLog@Utility]Error: " & ex.Message)
            Return "[AddToLog@Utility]Error: " & ex.Message
        End Try
    End Function
End Module
Module Memory
    Public ServerIP As String
    Public ServerPort As Integer
    Sub LoadMemory()
        Try
            Dim llaveReg As String = "SOFTWARE\\Zhenboro\\RMTMIC"
            Dim registerKey As RegistryKey = Registry.CurrentUser.OpenSubKey(llaveReg, True)
            If registerKey Is Nothing Then
                SaveMemory()
            Else
                ServerIP = registerKey.GetValue("ServerIP")
                ServerPort = registerKey.GetValue("ServerPort")
                Main.TextBox1.Text = ServerIP
                Main.TextBox2.Text = ServerPort
            End If
        Catch ex As Exception
            Console.WriteLine("LoadMemory Error: " & ex.Message)
        End Try
    End Sub
    Sub SaveMemory()
        Try
            Dim llaveReg As String = "SOFTWARE\\Zhenboro\\RMTMIC"
            Dim registerKey As RegistryKey = Registry.CurrentUser.OpenSubKey(llaveReg, True)
            If registerKey Is Nothing Then
                Registry.CurrentUser.CreateSubKey(llaveReg, True)
                registerKey = Registry.CurrentUser.OpenSubKey(llaveReg, True)
            End If
            If ServerIP = Nothing Or ServerPort = Nothing Then
                ServerIP = InputBox("Direccion IP", Main.Text)
                ServerPort = InputBox("Puerto", Main.Text)
            End If
            registerKey.SetValue("ServerIP", ServerIP)
            registerKey.SetValue("ServerPort", ServerPort)
            LoadMemory()
        Catch ex As Exception
            Console.WriteLine("SaveMemory Error: " & ex.Message)
        End Try
    End Sub
End Module
Module StartUp
    Sub Init()
        AddToLog("Init", My.Application.Info.AssemblyName & " " & My.Application.Info.Version.ToString & " (" & Application.ProductVersion & ")" & " has started! " & DateTime.Now.ToString("hh:mm:ss tt dd/MM/yyyy"), True)
        Try
            CommonActions()
            LoadMemory()
        Catch ex As Exception
            AddToLog("Init@StartUp", "Error: " & ex.Message, True)
        End Try
    End Sub
    Sub CommonActions()
        Try
            If Not My.Computer.FileSystem.DirectoryExists(DIRCommons) Then
                My.Computer.FileSystem.CreateDirectory(DIRCommons)
            End If
            If Not My.Computer.FileSystem.DirectoryExists(DIRHome) Then
                My.Computer.FileSystem.CreateDirectory(DIRHome)
            End If
        Catch ex As Exception
            AddToLog("CommonActions@StartUp", "Error: " & ex.Message, True)
        End Try
    End Sub
End Module