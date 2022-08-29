Public Class Keyboard
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal Hook As Integer, ByVal KeyDelegate As KDel, ByVal HMod As Integer, ByVal ThreadId As Integer) As Integer
    Private Declare Function CallNextHookEx Lib "user32" (ByVal Hook As Integer, ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KeyStructure) As Integer
    Private Declare Function UnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal Hook As Integer) As Integer
    Private Delegate Function KDel(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KeyStructure) As Integer
    Public Shared Event Down(ByVal Key As String)
    'Public Shared Event Up(ByVal Key As String)
    Private Shared Key As Integer
    Private Shared KHD As KDel
    Private Structure KeyStructure : Public Code As Integer : Public ScanCode As Integer : Public Flags As Integer : Public Time As Integer : Public ExtraInfo As Integer : End Structure
    Public Sub CreateHook()
        KHD = New KDel(AddressOf Proc)
        Key = SetWindowsHookEx(13, KHD, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
    End Sub

    Private Function Proc(ByVal Code As Integer, ByVal wParam As Integer, ByRef lParam As KeyStructure) As Integer
        If (Code = 0) Then
            Select Case wParam
                Case &H100, &H104 : RaiseEvent Down(lParam.Code.ToString)
                    'Case &H101, &H105 : RaiseEvent Up(Feed(CType(lParam.Code, Keys)))
            End Select
        End If
        Return CallNextHookEx(Key, Code, wParam, lParam)
    End Function
    Public Sub DiposeHook()
        UnhookWindowsHookEx(Key)
        MyBase.Finalize()
    End Sub
End Class
