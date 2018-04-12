Public Class frmAbout
    Sub New
        InitializeComponent()
    End Sub

    Public Overrides Sub ProcessCommand(ByVal cmd As System.Enum, ByVal arg As Object)
        MyBase.ProcessCommand(cmd, arg)
    End Sub

    Public Enum SplashScreenCommand
        SomeCommandId
    End Enum

    Private Sub HyperlinkLabelControl1_Click(sender As Object, e As EventArgs) Handles HyperlinkLabelControl1.Click
        Process.Start("mailto:walidpianooo@gmail.com")
    End Sub
End Class
