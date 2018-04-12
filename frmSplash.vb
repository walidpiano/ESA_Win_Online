Public Class frmSplash
    Sub New
        InitializeComponent()
    End Sub

    Public Overrides Sub ProcessCommand(ByVal cmd As System.Enum, ByVal arg As Object)
        MyBase.ProcessCommand(cmd, arg)
    End Sub

    Public Enum SplashScreenCommand
        SomeCommandId
    End Enum

    Private Sub frmSplash_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If ExClass.checkDBConnection Then
            ExClass.FillData()
        End If

        Me.Close()
        frmMain.Show()
    End Sub
End Class
