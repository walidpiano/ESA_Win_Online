Imports System.ComponentModel.DataAnnotations
Imports System.IO
Imports DevExpress.XtraLayout.Helpers
Imports DevExpress.XtraLayout
Imports DevExpress.XtraEditors
Imports System.Drawing.Imaging
Imports DevExpress.XtraBars.Docking2010.Views.WindowsUI
Imports DevExpress.XtraBars.Docking2010.Customization

Partial Public Class frmMain
    Dim SelectedMail As Int64 = 0
    Public Sub New()
        InitializeComponent()

    End Sub

    Public Sub Wait(ByVal wait As Boolean)
        If wait = True Then
            Try
                Me.SplashScreenManager1.ShowWaitForm()
            Catch ex As Exception

            End Try
        Else
            Try
                Me.SplashScreenManager1.CloseWaitForm()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub RemoveValidationNew(sender As Object, e As EventArgs)
        DxValidationProvider1.RemoveControlError(CType(sender, Windows.Forms.Control))
    End Sub

    Private Sub RemoveValidationOld(sender As Object, e As EventArgs)
        DxValidationProvider2.RemoveControlError(CType(sender, Windows.Forms.Control))
    End Sub

    Public Function ValidEntry(ByVal NewOrOld As Boolean) As Boolean
        Dim result As Boolean
        If NewOrOld Then
            result = DxValidationProvider1.Validate()
        Else
            result = DxValidationProvider2.Validate()
        End If

        Return result

    End Function

    Public Sub Clear(ByVal NorO As Boolean)
        If NorO Then
            pePhotoN.Image = Nothing
            luCateogryN.EditValue = Nothing
            luPlaceN.EditValue = Nothing
            luESAPointN.EditValue = Nothing
            teStudentN.Text = Nothing
            teTaxCodeN.Text = Nothing
            deDOBN.EditValue = Nothing
            mruNationalityN.EditValue = Nothing
            luCourseN.EditValue = Nothing
            luInstructorN.EditValue = Nothing
            deCourseDateN.EditValue = Nothing
            deRegDateN.EditValue = Nothing
            cbSexN.SelectedIndex = 0
            mruBirthPlaceN.EditValue = Nothing
            teHomePhoneN.Text = Nothing
            teCellPhoneN.Text = Nothing
            mruCountryN.EditValue = Nothing
            mruStateN.EditValue = Nothing
            mruCityN.EditValue = Nothing
            mruZIPCode.Text = Nothing
            teAddressN.Text = Nothing
            teEmailN.Text = Nothing
            meNotesN.Text = Nothing
            luCateogryN.Focus()
            ClearAllValidationNew()
        Else
            pePhotoO.Image = Nothing
            luCategoryO.EditValue = Nothing
            luPlaceO.EditValue = Nothing
            luESAPointO.EditValue = Nothing
            teStudentO.Text = Nothing
            luCourseO.EditValue = Nothing
            luInstructorO.EditValue = Nothing
            deCourseDateO.EditValue = Nothing
            meNotesO.Text = Nothing
            teESANumberO.Text = Nothing
            luCategoryO.Focus()
            ClearAllValidationOld()
        End If
        SelectedMail = 0

    End Sub



    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Application.Exit()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load


        If Not ExClass.checkDBConnection() Then
            ExClass.findDatabase()
            Exit Sub
        End If
        labelControl.BackColor = windowsUIButtonPanelMain.BackColor
        pbLogo.Parent = labelControl
        pbLogo.BackColor = Color.Transparent

        AddHandler pePhotoN.LostFocus, AddressOf RemoveValidationNew
        AddHandler luCateogryN.LostFocus, AddressOf RemoveValidationNew
        AddHandler luPlaceN.LostFocus, AddressOf RemoveValidationNew
        AddHandler luESAPointN.LostFocus, AddressOf RemoveValidationNew
        AddHandler luCourseN.LostFocus, AddressOf RemoveValidationNew
        AddHandler luInstructorN.LostFocus, AddressOf RemoveValidationNew
        AddHandler deCourseDateN.LostFocus, AddressOf RemoveValidationNew
        AddHandler teStudentN.LostFocus, AddressOf RemoveValidationNew
        AddHandler deRegDateN.LostFocus, AddressOf RemoveValidationNew
        AddHandler cbSexN.LostFocus, AddressOf RemoveValidationNew
        AddHandler deDOBN.LostFocus, AddressOf RemoveValidationNew
        AddHandler mruNationalityN.LostFocus, AddressOf RemoveValidationNew
        AddHandler teCellPhoneN.LostFocus, AddressOf RemoveValidationNew
        AddHandler mruCountryN.LostFocus, AddressOf RemoveValidationNew
        AddHandler mruStateN.LostFocus, AddressOf RemoveValidationNew
        AddHandler mruCityN.LostFocus, AddressOf RemoveValidationNew
        AddHandler teAddressN.LostFocus, AddressOf RemoveValidationNew
        AddHandler teEmailN.LostFocus, AddressOf RemoveValidationNew

        AddHandler luCategoryO.LostFocus, AddressOf RemoveValidationOld
        AddHandler luPlaceO.LostFocus, AddressOf RemoveValidationOld
        AddHandler luESAPointO.LostFocus, AddressOf RemoveValidationOld
        AddHandler luCourseO.LostFocus, AddressOf RemoveValidationOld
        AddHandler luInstructorO.LostFocus, AddressOf RemoveValidationOld
        AddHandler deCourseDateO.LostFocus, AddressOf RemoveValidationOld
        AddHandler teStudentO.LostFocus, AddressOf RemoveValidationOld

        ExClass.FillData()
        ExClass.FillNames(teStudentO)
        ExClass.FillNames(teStudentN)
        GridControl1.DataSource = Mail.getAllMails()
    End Sub



    Private Sub windowsUIButtonPanelMain_ButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.ButtonEventArgs) Handles windowsUIButtonPanelMain.ButtonClick
        If e.Button.Properties.Caption = "Clear" Or e.Button.Properties.Caption = "New" Then
            Clear(True)
            Clear(False)
        ElseIf e.Button.Properties.Caption = "Delete" Then
            If SelectedMail = 0 Then
                Exit Sub
            End If
            Dim diaR As DialogResult = MessageBox.Show("Are you sure you want to delete this registration?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If diaR = Windows.Forms.DialogResult.Yes Then
                Wait(True)
                Mail.Delete(SelectedMail)
                GridControl1.DataSource = Mail.getAllMails()
                Wait(False)
            End If
        ElseIf e.Button.Properties.Caption = "Save" Then
            Wait(True)
            Dim DataLine As String = ""
            If tsStatus.IsOn = False Then
                DataLine = ExClass.ConvertToLine(True, True)
            Else
                DataLine = ExClass.ConvertToLine(False, True)
            End If
            Dim dataLineList() As String
            dataLineList = Split(DataLine, ";")
            Dim subject As String = dataLineList(8)
            If subject = "" Then
                subject = "No Name"
            End If
            Dim newMail As New Mail(SelectedMail, subject, dataLineList(0), False, Now, DataLine)

            If SelectedMail = 0 Then
                newMail.Add()
            Else
                newMail.Update()
            End If

            GridControl1.DataSource = Mail.getAllMails
            Wait(False)
        ElseIf e.Button.Properties.Caption = "Send" Then
            Dim NewOrOld As Boolean
            If tsStatus.IsOn Then
                NewOrOld = False
            Else
                NewOrOld = True
            End If
            If ValidEntry(NewOrOld) Then
                Wait(True)
                Dim msg As String
                msg = ExSendMail.SendToESA
                If msg = "Registration has been successfully sent!" Then
                    Dim DataLine As String = ""
                    If tsStatus.IsOn = False Then
                        DataLine = ExClass.ConvertToLine(True, True)
                    Else
                        DataLine = ExClass.ConvertToLine(False, True)
                    End If
                    Dim dataLineList() As String
                    dataLineList = Split(DataLine, ";")
                    Dim newMail As New Mail(SelectedMail, dataLineList(8), dataLineList(0), True, Now, DataLine)
                    If SelectedMail = 0 Then
                        newMail.Add()
                    Else
                        newMail.Update()
                    End If
                    GridControl1.DataSource = Mail.getAllMails
                    Clear(True)
                    Clear(False)
                End If
                Wait(False)
                ExClass.showMessage("Registraion Send", msg)
                
            End If
        ElseIf e.Button.Properties.Caption = "Sync" Then

            Wait(True)
            ExSync.Synchronize()
            ExClass.FillData()
            ExClass.FillNames(teStudentO)
            ExClass.FillNames(teStudentN)

            Wait(False)
        ElseIf e.Button.Properties.Caption = "DB" Then
            ExClass.findDatabase()
        ElseIf e.Button.Properties.Caption = "About" Then
            frmAbout.Opacity = 100%
            frmAbout.ShowDialog()
        End If

    End Sub

    Private Sub ClearAllValidationNew()
        DxValidationProvider1.RemoveControlError(pePhotoN)
        DxValidationProvider1.RemoveControlError(luCateogryN)
        DxValidationProvider1.RemoveControlError(luPlaceN)
        DxValidationProvider1.RemoveControlError(luESAPointN)
        DxValidationProvider1.RemoveControlError(luCourseN)
        DxValidationProvider1.RemoveControlError(luInstructorN)
        DxValidationProvider1.RemoveControlError(deCourseDateN)
        DxValidationProvider1.RemoveControlError(teStudentN)
        DxValidationProvider1.RemoveControlError(deRegDateN)
        DxValidationProvider1.RemoveControlError(cbSexN)
        DxValidationProvider1.RemoveControlError(deDOBN)
        DxValidationProvider1.RemoveControlError(mruNationalityN)
        DxValidationProvider1.RemoveControlError(teCellPhoneN)
        DxValidationProvider1.RemoveControlError(mruCountryN)
        DxValidationProvider1.RemoveControlError(mruStateN)
        DxValidationProvider1.RemoveControlError(mruCityN)
        DxValidationProvider1.RemoveControlError(teAddressN)
        DxValidationProvider1.RemoveControlError(teEmailN)
    End Sub

    Private Sub ClearAllValidationOld()
        DxValidationProvider2.RemoveControlError(pePhotoO)
        DxValidationProvider2.RemoveControlError(luCategoryO)
        DxValidationProvider2.RemoveControlError(luPlaceO)
        DxValidationProvider2.RemoveControlError(luESAPointO)
        DxValidationProvider2.RemoveControlError(luCourseO)
        DxValidationProvider2.RemoveControlError(luInstructorO)
        DxValidationProvider2.RemoveControlError(deCourseDateO)
        DxValidationProvider2.RemoveControlError(teStudentO)
        DxValidationProvider2.RemoveControlError(deCourseDateO)
    End Sub

    Private Sub luCategoryO_EditValueChanged(sender As Object, e As EventArgs) Handles luCategoryO.EditValueChanged
        ExClass.FillCourses(CInt(luCategoryO.EditValue), luCourseO)
    End Sub

    Private Sub luCategoryN_EditValueChanged(sender As Object, e As EventArgs) Handles luCateogryN.EditValueChanged
        ExClass.FillCourses(CInt(luCateogryN.EditValue), luCourseN)
    End Sub

    Private Sub luCourseO_EditValueChanged(sender As Object, e As EventArgs) Handles luCourseO.EditValueChanged
        ExClass.FillInstructors(CInt(luCourseO.EditValue), luInstructorO)
    End Sub

    Private Sub luCourseN_EditValueChanged(sender As Object, e As EventArgs) Handles luCourseN.EditValueChanged
        ExClass.FillInstructors(CInt(luCourseN.EditValue), luInstructorN)
    End Sub

    Private Sub mruCountryN_EditValueChanged(sender As Object, e As EventArgs) Handles mruCountryN.EditValueChanged
        ExClass.FillState(mruCountryN.Text, mruStateN)
    End Sub

    Private Sub mruStateN_EditValueChanged(sender As Object, e As EventArgs) Handles mruStateN.EditValueChanged
        ExClass.FillCity(mruStateN.Text, mruCityN)
    End Sub

    Private Sub pePhotoN_DoubleClick(sender As Object, e As EventArgs) Handles pePhotoN.DoubleClick
        Me.pePhotoN.LoadImage()
    End Sub

    Private Sub pePhotoO_DoubleClick(sender As Object, e As EventArgs) Handles pePhotoO.DoubleClick
        Me.pePhotoO.LoadImage()
    End Sub

    Private Sub teCellPhoneN_LostFocus(sender As Object, e As EventArgs) Handles teCellPhoneN.LostFocus
        DxValidationProvider1.RemoveControlError(teCellPhoneN)
    End Sub


    Private Sub tsStatus_Toggled(sender As Object, e As EventArgs) Handles tsStatus.Toggled
        If tsStatus.IsOn Then
            NavigationFrame1.SelectedPageIndex = 1
        Else
            NavigationFrame1.SelectedPageIndex = 0
        End If
    End Sub

    Private Sub mruCityN_EditValueChanged(sender As Object, e As EventArgs) Handles mruCityN.EditValueChanged
        ExClass.FillZIPCodes(mruCityN.Text, mruZIPCode)
    End Sub

    Private Sub GridView1_RowClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles GridView1.RowClick
        Dim id As Int64
        id = GridView1.GetFocusedRowCellValue("ID")

        ExClass.recallMail(id)
        SelectedMail = id
    End Sub

    
    Private Sub GridView1_SelectionChanged(sender As Object, e As DevExpress.Data.SelectionChangedEventArgs) Handles GridView1.SelectionChanged
        Dim id As Int64
        id = GridView1.GetFocusedRowCellValue("ID")

        ExClass.recallMail(id)
        SelectedMail = id
    End Sub

    'Private Sub test()
    '    Dim button As New SimpleButton() With {.Text = "ShowFlyout"}
    '    button.Dock = DockStyle.Top
    '    button.Parent = NavigationPage1
    '    Dim action As New FlyoutAction()
    '    action.Caption = "Flyout Action"
    '    action.Description = "Flyout Action Description"
    '    action.Commands.Add(FlyoutCommand.OK)
    '    AddHandler button.Click, Sub(sender, e) FlyoutDialog.Show(NavigationPage1.FindForm(), action)
    'End Sub
End Class
