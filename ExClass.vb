Imports System.IO
Imports System.Drawing.Imaging
Imports System.Data.SqlClient
Imports ESA_Online.Store
Imports DevExpress.XtraBars.Docking2010.Views.WindowsUI
Imports DevExpress.XtraBars.Docking2010.Customization
Imports System.Security.AccessControl


Public Class ExClass
    Public Shared myLocalConn As New SqlConnection(GV.myLocalConn)
    Public Shared myOnlineConn As New SqlConnection(GV.myOnlineConn)
    Public Shared Function Base64ToImage(ByVal base64Image As String) As Image
        Using ms As New MemoryStream(Convert.FromBase64String(base64Image))
            Dim image As Image = image.FromStream(ms, True)
            Return image
        End Using
    End Function

    Public Shared Sub toConnect(ByVal Connect As Boolean, ByVal Local As Boolean)
        If Local = True Then
            If Connect = True Then
                If myLocalConn.State = ConnectionState.Closed Then
                    myLocalConn.Open()
                End If
            Else
                myLocalConn.Close()
            End If
        Else
            If Connect = True Then
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
            Else
                myOnlineConn.Close()
            End If
        End If

    End Sub

    Public Shared Function ImageToBase64(ByVal Image As Image) As String
        Dim Result As String

        Dim newImage As Image
        newImage = LowerImage(Image)
        Using ms As New MemoryStream()
            Image.Save(ms, ImageFormat.Png)
            Dim imageBytes As Byte() = ms.ToArray()
            Dim Base64String As String = Convert.ToBase64String(imageBytes)
            Result = Base64String
        End Using
        Return Result
    End Function

    Public Shared Sub FillCourseCategory(ByVal lue As DevExpress.XtraEditors.LookUpEdit)
        Dim Query As String = "SELECT * FROM tblCourseCategories ORDER BY Category;"

        Using cmd = New SqlCommand(Query, myLocalConn)
            toConnect(True, True)
            Dim adt As New SqlDataAdapter
            Dim ds As New DataSet()
            adt.SelectCommand = cmd
            adt.Fill(ds)
            adt.Dispose()

            lue.Properties.DataSource = ds.Tables(0)
            lue.Properties.DisplayMember = "Category"
            lue.Properties.ValueMember = "ID"
            lue.EditValue = Nothing

            myLocalConn.Close()

        End Using
    End Sub

    Public Shared Sub FillCourses(ByVal Category As Integer, ByVal Lue As DevExpress.XtraEditors.LookUpEdit)

        If Category = Nothing Then
            Lue.Properties.DataSource = Nothing
            Lue.Properties.DisplayMember = Nothing
            Lue.Properties.ValueMember = Nothing
            Exit Sub
        End If

        Dim Query As String
        Query = "SELECT ID, CourseName FROM tblCourses WHERE CourseCategory = @Category ORDER BY CourseName;"

        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@Category", Category)
            toConnect(True, True)
            Dim adt As New SqlDataAdapter
            Dim ds As New DataSet()
            adt.SelectCommand = cmd
            adt.Fill(ds)
            adt.Dispose()

            Lue.Properties.DataSource = Nothing
            Lue.Properties.DataSource = ds.Tables(0)
            Lue.Properties.DisplayMember = "CourseName"
            Lue.Properties.ValueMember = "ID"
            Lue.EditValue = Nothing

            myLocalConn.Close()

        End Using
    End Sub

    Public Shared Sub FillCoursePlaces(ByVal Lue As DevExpress.XtraEditors.LookUpEdit)
        Dim Query As String = "SELECT * FROM tblCoursePlaces ORDER BY CoursePlace;"
        Using cmd = New SqlCommand(Query, myLocalConn)
            toConnect(True, True)
            Dim adt As New SqlDataAdapter
            Dim ds As New DataSet()
            adt.SelectCommand = cmd
            adt.Fill(ds)
            adt.Dispose()

            Lue.Properties.DataSource = ds.Tables(0)
            Lue.Properties.DisplayMember = "CoursePlace"
            Lue.Properties.ValueMember = "ID"
            Lue.EditValue = Nothing

            myLocalConn.Close()

        End Using
    End Sub

    Public Shared Sub FillESAPoint(ByVal Lue As DevExpress.XtraEditors.LookUpEdit)
        Dim Query As String = "SELECT ID, PointName FROM tblESAPoint ORDER BY PointName"
        Using cmd = New SqlCommand(Query, myLocalConn)
            toConnect(True, True)
            Dim adt As New SqlDataAdapter
            Dim ds As New DataSet()
            adt.SelectCommand = cmd
            adt.Fill(ds)
            adt.Dispose()

            Lue.Properties.DataSource = ds.Tables(0)
            Lue.Properties.DisplayMember = "PointName"
            Lue.Properties.ValueMember = "ID"
            Lue.EditValue = Nothing
            myLocalConn.Close()

        End Using
    End Sub

    Public Shared Sub FillInstructors(ByVal Course As Integer, ByVal Lue As DevExpress.XtraEditors.LookUpEdit)
        Dim Query As String = ""
        If Course = 0 Then
            Query = "SELECT ID, Name FROM tblInstructors ORDER BY Name;"
        Else
            Query = "SELECT tblInstructors.ID, tblInstructors.Name" _
                   & " FROM tblInstructors INNER JOIN tblInstructorCourses ON  tblInstructors.ID = tblInstructorCourses.Instructor" _
                   & " WHERE tblInstructorCourses.Course = @Course" _
                   & " ORDER BY tblInstructors.Name;"
        End If

        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@Course", Course)
            toConnect(True, True)
            Dim adt As New SqlDataAdapter
            Dim ds As New DataSet()
            adt.SelectCommand = cmd
            adt.Fill(ds)
            adt.Dispose()

            Lue.Properties.DataSource = ds.Tables(0)
            Lue.Properties.DisplayMember = "Name"
            Lue.Properties.ValueMember = "ID"
            Lue.EditValue = Nothing

            myLocalConn.Close()

        End Using
    End Sub

    Public Shared Sub FillCountry(ByVal Mru As DevExpress.XtraEditors.MRUEdit)
        Dim Query As String = "SELECT Country FROM tblCountry ORDER BY Country;"
        Dim dt As New DataTable

        Using cmd = New SqlCommand(Query, myLocalConn)
            toConnect(True, True)
            Using dr As SqlDataReader = cmd.ExecuteReader
                dt.Load(dr)
            End Using
            myLocalConn.Close()
        End Using

        Dim Countries(dt.Rows.Count - 1) As String
        For x As Integer = 0 To dt.Rows.Count - 1
            Countries(x) = dt.Rows(x)(0).ToString
        Next

        Mru.Properties.Items.Clear()
        Mru.Properties.Items.AddRange(Countries)
        Mru.SelectedText = Nothing
        Mru.Text = Nothing
    End Sub

    Public Shared Sub FillAllStates(ByVal Mru As DevExpress.XtraEditors.MRUEdit)
        Dim Query As String = "SELECT StateName FROM tblState ORDER BY StateName;"
        Dim dt As New DataTable
        Using cmd = New SqlCommand(Query, myLocalConn)
            toConnect(True, True)
            Using dr As SqlDataReader = cmd.ExecuteReader
                dt.Load(dr)
            End Using
            myLocalConn.Close()
        End Using

        Dim States(dt.Rows.Count - 1) As String
        For x As Integer = 0 To dt.Rows.Count - 1
            States(x) = dt.Rows(x)(0).ToString
        Next

        Mru.Properties.Items.Clear()
        Mru.Properties.Items.AddRange(States)
        Mru.SelectedText = Nothing
        Mru.Text = Nothing

    End Sub

    Public Shared Sub FillState(ByVal Country As String, ByVal Mru As DevExpress.XtraEditors.MRUEdit)
        Dim Query As String = "SELECT tblState.StateName" _
                              & "         FROM tblState" _
                              & " INNER JOIN tblCountry ON tblState.Country = tblCountry.ID" _
                              & " WHERE tblCountry.Country = @Country ORDER BY tblState.StateName;"
        Dim dt As New DataTable

        Using cmd = New SqlCommand(Query, myLocalConn)
            toConnect(True, True)
            cmd.Parameters.AddWithValue("@Country", Country)
            Using dr As SqlDataReader = cmd.ExecuteReader
                dt.Load(dr)
            End Using
            myLocalConn.Close()
        End Using

        Dim States(dt.Rows.Count - 1) As String
        For x As Integer = 0 To dt.Rows.Count - 1
            States(x) = dt.Rows(x)(0).ToString
        Next

        Mru.Properties.Items.Clear()
        Mru.Properties.Items.AddRange(States)
        Mru.SelectedText = Nothing
        Mru.Text = Nothing

    End Sub

    Public Shared Sub FillCity(ByVal State As String, ByVal Mru As DevExpress.XtraEditors.MRUEdit)
        Dim Query As String = "SELECT tblCity.CityName" _
                              & "         FROM tblCity" _
                              & " INNER JOIN tblState ON tblCity.StateID = tblState.ID" _
                              & " WHERE tblState.StateName = @StateName" _
                              & " ORDER BY tblCity.CityName;"
        Dim dt As New DataTable

        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@StateName", State)
            toConnect(True, True)
            Using dr As SqlDataReader = cmd.ExecuteReader
                dt.Load(dr)
            End Using
            myLocalConn.Close()
        End Using

        Dim Cities(dt.Rows.Count - 1) As String
        For x As Integer = 0 To dt.Rows.Count - 1
            Cities(x) = dt.Rows(x)(0).ToString
        Next

        Mru.Properties.Items.Clear()
        Mru.Properties.Items.AddRange(Cities)
        Mru.SelectedText = Nothing
        Mru.Text = Nothing

    End Sub

    Public Shared Sub FillZIPCodes(ByVal City As String, ByVal Mru As DevExpress.XtraEditors.MRUEdit)
        Dim Query As String = "SELECT ZIPCode" _
                              & " FROM tblCity" _
                              & " WHERE CityName = @CityName" _
                              & " ORDER BY ZIPCode;"

        Dim dt As New DataTable
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@CityName", City)
            toConnect(True, True)
            Using dr As SqlDataReader = cmd.ExecuteReader
                dt.Load(dr)
            End Using
            myLocalConn.Close()
        End Using

        Dim ZIPCodes(dt.Rows.Count - 1) As String
        For x As Integer = 0 To dt.Rows.Count - 1
            ZIPCodes(x) = dt.Rows(x)(0).ToString
        Next

        Mru.Properties.Items.Clear()
        Mru.Properties.Items.AddRange(ZIPCodes)
        Mru.SelectedText = Nothing
        Mru.Text = Nothing

    End Sub

    Public Shared Sub FillData()
        If checkDBConnection() Then
            FillCourseCategory(frmMain.luCateogryN)
            FillCoursePlaces(frmMain.luPlaceN)
            FillESAPoint(frmMain.luESAPointN)
            FillCountry(frmMain.mruNationalityN)
            FillCountry(frmMain.mruCountryN)
            FillAllStates(frmMain.mruBirthPlaceN)
            FillCourseCategory(frmMain.luCategoryO)
            FillCoursePlaces(frmMain.luPlaceO)
            FillESAPoint(frmMain.luESAPointO)
        End If
    End Sub

    Public Shared Function ConvertToLine(ByVal NewOrOld As Boolean, ByVal WithPhoto As Boolean) As String
        Dim CashLine As String = ""

        Dim Status As String = ""
        Dim Photo As String = ""
        Dim Category As String = ""
        Dim Place As String = ""
        Dim Point As String = ""
        Dim Course As String = ""
        Dim Instructor As String = ""
        Dim CourseDate As String = ""
        Dim Student As String = ""
        Dim ESANumber As String = ""
        Dim TaxCode As String = ""
        Dim DOB As String = ""
        Dim Nationality As String = ""
        Dim RegDate As String = ""
        Dim Sex As String = ""
        Dim BirthPlace As String = ""
        Dim HomePhone As String = ""
        Dim CellPhone As String = ""
        Dim Country As String = ""
        Dim State As String = ""
        Dim City As String = ""
        Dim ZIP As String = ""
        Dim Address As String = ""
        Dim Email As String = ""
        Dim Notes As String = ""


        If NewOrOld Then
            Status = "NEW"
            With frmMain
                If Not .pePhotoN.Image Is Nothing AndAlso WithPhoto Then
                    Photo = ImageToBase64(.pePhotoN.Image)
                End If
                Category = .luCateogryN.EditValue
                Place = .luPlaceN.EditValue
                Point = .luESAPointN.EditValue
                Course = .luCourseN.EditValue
                Instructor = .luInstructorN.EditValue
                CourseDate = (CDate(.deCourseDateN.EditValue).Ticks)
                Student = StrConv(.teStudentN.Text, VbStrConv.ProperCase)
                ESANumber = ""
                TaxCode = .teTaxCodeN.Text.ToUpper
                DOB = (CDate(.deDOBN.EditValue).Ticks)
                Nationality = StrConv(.mruNationalityN.Text, VbStrConv.ProperCase)
                RegDate = (CDate(.deRegDateN.EditValue).Ticks)
                Sex = .cbSexN.SelectedIndex
                BirthPlace = StrConv(.mruBirthPlaceN.Text, VbStrConv.ProperCase)
                HomePhone = .teHomePhoneN.Text
                CellPhone = .teCellPhoneN.Text
                Country = StrConv(.mruCountryN.Text, VbStrConv.ProperCase)
                State = StrConv(.mruStateN.Text, VbStrConv.ProperCase)
                City = StrConv(.mruCityN.Text, VbStrConv.ProperCase)
                ZIP = .mruZIPCode.Text
                Address = .teAddressN.Text
                Email = .teEmailN.Text
                Notes = .meNotesN.Text.Replace(";", " ")
            End With
        Else
            Status = "Old"
            With frmMain
                If Not .pePhotoO.Image Is Nothing AndAlso WithPhoto Then
                    Photo = ImageToBase64(.pePhotoO.Image)
                End If
                Category = .luCategoryO.EditValue
                Place = .luPlaceO.EditValue
                Point = .luESAPointO.EditValue
                Course = .luCourseO.EditValue
                Instructor = .luInstructorO.EditValue
                CourseDate = (CDate(.deCourseDateO.EditValue).Ticks).ToString
                Student = StrConv(.teStudentO.Text, VbStrConv.ProperCase)
                ESANumber = .teESANumberO.Text.ToUpper
                Notes = .meNotesO.Text.Replace(";", " ")
            End With
        End If
        CashLine = Status & ";" & Photo & ";" & Category & ";" & Place & ";" & Point & ";" & Course & ";" & Instructor _
            & ";" & CourseDate & ";" & Student & ";" & ESANumber & ";" & TaxCode & ";" & DOB & ";" & Nationality.Replace(";", "") & ";" _
            & RegDate & ";" & Sex & ";" & BirthPlace.Replace(";", "") & ";" & HomePhone & ";" & CellPhone & ";" & Country.Replace(";", "") & ";" & State.Replace(";", "") & ";" _
            & City.Replace(";", "") & ";" & ZIP.Replace(";", "") & ";" & Address & ";" & Email & ";" & Notes & ";"

        Return CashLine
    End Function

    Public Shared Sub recallMail(ByVal ID As Int64)
        'Dim DataLine As String = My.Settings.Registration
        'If DataLine = "" Then
        'Exit Sub
        'End If
        If ID = 0 Then
            Exit Sub
        End If
        Dim DataLine As String = Mail.Recall(ID)

        Dim Status As String = ""
        Dim Photo As String = ""
        Dim Category As String = ""
        Dim Place As String = ""
        Dim Point As String = ""
        Dim Course As String = ""
        Dim Instructor As String = ""
        Dim CourseDate As String = ""
        Dim Student As String = ""
        Dim ESANumber As String = ""
        Dim TaxCode As String = ""
        Dim DOB As String = ""
        Dim Nationality As String = ""
        Dim RegDate As String = ""
        Dim Sex As String = ""
        Dim BirthPlace As String = ""
        Dim HomePhone As String = ""
        Dim CellPhone As String = ""
        Dim Country As String = ""
        Dim State As String = ""
        Dim City As String = ""
        Dim ZIP As String = ""
        Dim Address As String = ""
        Dim Email As String = ""
        Dim Notes As String = ""

        Dim EntData() As String
        EntData = DataLine.Split(CChar(";"))

        Photo = EntData(1)
        Category = EntData(2)
        Place = EntData(3)
        Point = EntData(4)
        Course = EntData(5)
        Instructor = EntData(6)
        CourseDate = EntData(7)
        Student = EntData(8)
        ESANumber = EntData(9)
        TaxCode = EntData(10)
        DOB = EntData(11)
        Nationality = EntData(12)
        RegDate = EntData(13)
        Sex = EntData(14)
        BirthPlace = EntData(15)
        HomePhone = EntData(16)
        CellPhone = EntData(17)
        Country = EntData(18)
        State = EntData(19)
        City = EntData(20)
        ZIP = EntData(21)
        Address = EntData(22)
        Email = EntData(23)
        Notes = EntData(24)


        Status = EntData(0)


        If EntData(0) = "NEW" Then
            With frmMain
                .tsStatus.IsOn = False
                If Photo = "" Then
                    .pePhotoN.Image = Nothing
                Else
                    .pePhotoN.Image = Base64ToImage(Photo)
                End If
                .luCateogryN.EditValue = .luCateogryN.Properties.GetKeyValueByDisplayText(GetCategory(CInt(Val(Category))))
                .luPlaceN.EditValue = .luPlaceN.Properties.GetKeyValueByDisplayText(GetPlace(CInt(Val(Place))))
                .luESAPointN.EditValue = .luESAPointN.Properties.GetKeyValueByDisplayText(GetPoint(CInt(Val(Point))))
                .luCourseN.EditValue = .luCourseN.Properties.GetKeyValueByDisplayText(GetCourse(CInt(Val(Course))))
                .luInstructorN.EditValue = .luInstructorN.Properties.GetKeyValueByDisplayText(GetInstructor(CInt(Val(Instructor))))
                .deCourseDateN.EditValue = New Date(Convert.ToInt64(CourseDate))
                .teStudentN.Text = Student
                .teESANumberO.Text = ""
                .teTaxCodeN.Text = TaxCode
                .deDOBN.EditValue = New Date(Convert.ToInt64(Val(DOB)))
                .mruNationalityN.Text = Nationality
                .deRegDateN.EditValue = New Date(Convert.ToInt64(Val(RegDate)))
                .cbSexN.SelectedIndex = CInt(Val(Sex))
                .mruBirthPlaceN.Text = BirthPlace
                .teHomePhoneN.Text = HomePhone
                .teCellPhoneN.Text = CellPhone
                .mruCountryN.Text = Country
                .mruStateN.Text = State
                .mruCityN.Text = City
                .mruZIPCode.Text = ZIP
                .teAddressN.Text = Address
                .teEmailN.Text = Email
                .meNotesN.Text = Notes

                frmMain.DxValidationProvider1.Validate()

            End With
        Else
            With frmMain
                .tsStatus.IsOn = True
                If Photo = "" Then
                    .pePhotoO.Image = Nothing
                Else
                    .pePhotoO.Image = Base64ToImage(Photo)
                End If
                .luCategoryO.EditValue = .luCategoryO.Properties.GetKeyValueByDisplayText(GetCategory(CInt(Val(Category))))
                .luPlaceO.EditValue = .luPlaceO.Properties.GetKeyValueByDisplayText(GetPlace(CInt(Val(Place))))
                .luESAPointO.EditValue = .luESAPointO.Properties.GetKeyValueByDisplayText(GetPoint(CInt(Val(Point))))
                .luCourseO.EditValue = .luCourseO.Properties.GetKeyValueByDisplayText(GetCourse(CInt(Val(Course))))
                .luInstructorO.EditValue = .luInstructorO.Properties.GetKeyValueByDisplayText(GetInstructor(CInt(Val(Instructor))))
                .deCourseDateO.EditValue = New Date(Convert.ToInt64(CourseDate))
                .teStudentO.Text = Student
                .teESANumberO.Text = ESANumber
                .meNotesO.Text = Notes

                frmMain.DxValidationProvider2.Validate()


            End With
        End If

    End Sub

    Public Shared Function GetPlace(ByVal ID As Long) As String
        Dim Query As String = "SELECT CoursePlace FROM tblCoursePlaces WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function

    Public Shared Function GetPoint(ByVal ID As Long) As String
        Dim Query As String = "SELECT PointName FROM tblESAPoint WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function

    Public Shared Function GetCategory(ByVal ID As Long) As String
        Dim Query As String = "SELECT Category FROM tblCourseCategories WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function
    Public Shared Function GetCountry(ByVal ID As Long) As String
        Dim Query As String = "SELECT Country FROM tblCountry WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function

    Public Shared Function GetState(ByVal ID As Long) As String
        Dim Query As String = "SELECT StateName FROM tblState WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function

    Public Shared Function GetCity(ByVal ID As Long) As String
        Dim Query As String = "SELECT CityName FROM tblCity WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function

    Public Shared Function GetCourse(ByVal ID As Long) As String
        Dim Query As String = "SELECT CourseName FROM tblCourses WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function

    Public Shared Function GetInstructor(ByVal ID As Long) As String
        Dim Query As String = "SELECT Name FROM tblInstructors WHERE ID = @ID;"
        Dim Result As String = ""
        Using cmd = New SqlCommand(Query, myLocalConn)
            cmd.Parameters.AddWithValue("@ID", ID)
            toConnect(True, True)
            Using dt As SqlDataReader = cmd.ExecuteReader
                If dt.Read() Then
                    Result = dt(0).ToString
                End If
            End Using
            myLocalConn.Close()
        End Using
        Return Result
    End Function

    Public Shared Function LowerImage(ByVal Image As Image) As Image
        Dim imgThumb As Image = Nothing

        Dim oWidth As Integer = Image.Width
        Dim oHeight As Integer = Image.Height

        Dim nWidth, nHeight As Integer
        nWidth = 400
        nHeight = CInt(Val(Math.Round((nWidth * oHeight) / oWidth, 0, MidpointRounding.AwayFromZero)))

        imgThumb = Image.GetThumbnailImage(nWidth, nHeight, Nothing, New IntPtr(0))

        Return imgThumb

    End Function

    Public Shared Sub FillNames(ByVal Mru As DevExpress.XtraEditors.MRUEdit)
        Dim Query As String = "SELECT DISTINCT [Subject] FROM tblMails ORDER BY [Subject];"

        Dim dt As New DataTable
        Using cmd = New SqlCommand(Query, myLocalConn)
            toConnect(True, True)
            Using dr As SqlDataReader = cmd.ExecuteReader
                dt.Load(dr)
            End Using
            myLocalConn.Close()
        End Using

        Dim Names(dt.Rows.Count - 1) As String
        For x As Integer = 0 To dt.Rows.Count - 1
            Names(x) = dt.Rows(x)(0).ToString
        Next

        Mru.Properties.Items.Clear()
        Mru.Properties.Items.AddRange(Names)
        Mru.SelectedText = Nothing
        Mru.Text = Nothing

    End Sub

    Public Shared Sub showMessage(ByVal Caption As String, ByVal Description As String)
        Dim action As New FlyoutAction()
        action.Caption = Caption
        action.Description = Description
        action.Commands.Add(FlyoutCommand.OK)
        FlyoutDialog.Show(frmMain.NavigationPage1.FindForm(), action)
    End Sub

    Public Shared Sub findDatabase()
        Dim openFD As OpenFileDialog = New OpenFileDialog()
        openFD.Filter = "Database (*.mdf)|*.mdf;"
        openFD.Title = "Select ESA_Online database"

        Dim dbPath As String
        Dim dbLogPath As String
        If openFD.ShowDialog() = DialogResult.OK Then
            dbPath = openFD.FileName
            dbLogPath = dbPath.Replace(".mdf", "_log.ldf")
            My.Settings.DatabasePath = dbPath
            If GrantAccess(dbPath) AndAlso GrantAccess(dbLogPath) Then
                My.Settings.Save()
                If checkDBConnection() Then
                    showMessage("Database Connection", "ESA_Online is connected to the database successfully. Please restart the application!")
                Else
                    showMessage("Database Connection", "Please restart the application, to make sure that the database is connected properly!")
                    frmSplash.Close()
                    frmMain.Close()
                End If

            Else
                showMessage("Database Connection", "ESA_Online couldn't establish the connection! Please run ESA_Online As Administrator, and try again!")
                frmSplash.Close()
                frmMain.Close()
            End If
        Else
            If Not checkDBConnection() Then
                frmSplash.Close()
                frmMain.Close()
            End If
        End If
    End Sub

    Public Shared Function GrantAccess(ByVal fullPath As String) As Boolean
        Dim dInfo As DirectoryInfo = New DirectoryInfo(fullPath)
        Dim dSecurity As DirectorySecurity = dInfo.GetAccessControl()
        dSecurity.AddAccessRule(New FileSystemAccessRule("everyone", FileSystemRights.FullControl, InheritanceFlags.ObjectInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow))
        Try
            dInfo.SetAccessControl(dSecurity)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Shared Function checkDBConnection() As Boolean
        Dim Query As String = "SELECT TOP(1)Connected FROM tblCheck;"
        Dim result As Boolean
        If My.Settings.DatabasePath = "" Then
            Return False
        End If
        Using cmd = New SqlCommand(Query, myLocalConn)
            Try
                If Not myLocalConn.State = ConnectionState.Open Then
                    myLocalConn.Open()
                End If
                result = cmd.ExecuteScalar
                myLocalConn.Close()
            Catch ex As Exception
                result = False
            End Try
        End Using
        
        Return result
    End Function
End Class
