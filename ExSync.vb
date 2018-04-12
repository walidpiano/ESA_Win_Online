Imports System.Data.SqlClient

Public Class ExSync
    Public Shared myLocalConn As New SqlConnection(GV.myLocalConn)
    Public Shared myOnlineConn As New SqlConnection(GV.myOnlineConn)

    Private Shared Sub ShrinkDB(ByVal DatabaseName As String)
        Dim Query As String = "BEGIN TRY" _
                              & " 	DBCC SHRINKDATABASE ('" & DatabaseName & "')" _
                              & " END TRY" _
                              & " BEGIN CATCH" _
                              & " END CATCH"
        Using cmd = New SqlCommand(Query, myLocalConn)
            Try
                If myLocalConn.State = ConnectionState.Closed Then
                    myLocalConn.Open()
                End If
                cmd.ExecuteNonQuery()
                myLocalConn.Close()
            Catch ex As Exception

            End Try
        End Using
    End Sub
    Private Shared Sub SyncCourseCategories()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, REPLACE(Category, '''', '') AS Category FROM tblCourseCategories"
        Dim Query2 As String = "DELETE FROM tblCourseCategories; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try
        

        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        Query2 &= "INSERT INTO tblCourseCategories (ID, Category) VALUES "
        For x As Integer = 0 To dt.Rows.Count - 1
            Query2 &= "(" & dt.Rows(x)(0).ToString & ", '" & dt.Rows(x)(1).ToString & "'), "
        Next
        Query2 = Query2.Substring(0, Query2.Length - 2)
        Query2 &= ";"

        Using cmd = New SqlCommand(Query2, myLocalConn)
            If myLocalConn.State = ConnectionState.Closed Then
                myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            myLocalConn.Close()
        End Using

    End Sub
    Private Shared Sub SyncInstructorCourses()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT Instructor, Course FROM tblInstructorCourses"
        Dim Query2 As String = "DELETE FROM tblInstructorCourses; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        Dim rCount As Integer = dt.Rows.Count
        If rCount = 0 Then
            Exit Sub
        End If

        For y As Integer = 0 To rCount - 1 Step 1000
            Query2 &= "INSERT INTO tblInstructorCourses (Instructor, Course) VALUES "
            For x As Integer = 0 To 999
                If x + y >= rCount Then
                    Exit For
                End If
                Query2 &= "(" & dt.Rows(x + y)(0).ToString & ", '" & dt.Rows(x + y)(1).ToString & "'), "

            Next
            Query2 = Query2.Substring(0, Query2.Length - 2)
            Query2 &= ";"


            Using cmd = New SqlCommand(Query2, myLocalConn)
                If myLocalConn.State = ConnectionState.Closed Then
                    myLocalConn.Open()
                End If
                cmd.ExecuteNonQuery()
                myLocalConn.Close()
            End Using
            
            Query2 = ""
        Next y

    End Sub
    Private Shared Sub SyncCoursePlaces()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, REPLACE(CoursePlace, '''', '') AS CoursePlace FROM tblCoursePlaces"
        Dim Query2 As String = "DELETE FROM tblCoursePlaces; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        Query2 &= "INSERT INTO tblCoursePlaces (ID, CoursePlace) VALUES "
        For x As Integer = 0 To dt.Rows.Count - 1
            Query2 &= "(" & dt.Rows(x)(0).ToString & ", '" & dt.Rows(x)(1).ToString & "'), "
        Next
        Query2 = Query2.Substring(0, Query2.Length - 2)
        Query2 &= ";"

        Using cmd = New SqlCommand(Query2, myLocalConn)
            If myLocalConn.State = ConnectionState.Closed Then
                myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            myLocalConn.Close()
        End Using

    End Sub
    Private Shared Sub SyncCountry()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, REPLACE(Country, '''', '') AS Country FROM tblCountry"
        Dim Query2 As String = "DELETE FROM tblCountry; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        Query2 &= "INSERT INTO tblCountry (ID, Country) VALUES "
        For x As Integer = 0 To dt.Rows.Count - 1
            Query2 &= "(" & dt.Rows(x)(0).ToString & ", '" & dt.Rows(x)(1).ToString & "'), "
        Next
        Query2 = Query2.Substring(0, Query2.Length - 2)
        Query2 &= ";"

        Using cmd = New SqlCommand(Query2, myLocalConn)
            If myLocalConn.State = ConnectionState.Closed Then
                myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            myLocalConn.Close()
        End Using
        

    End Sub

    Private Shared Sub SyncState()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, Country, REPLACE(StateName, '''', '') AS StateName FROM tblState"
        Dim Query2 As String = "DELETE FROM tblState; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        Query2 &= "INSERT INTO tblState (ID, Country, StateName) VALUES "
        For x As Integer = 0 To dt.Rows.Count - 1
            Query2 &= "(" & dt.Rows(x)(0).ToString & ", " & dt.Rows(x)(1).ToString & ", '" & dt.Rows(x)(2).ToString & "'), "
        Next
        Query2 = Query2.Substring(0, Query2.Length - 2)
        Query2 &= ";"

        Using cmd = New SqlCommand(Query2, myLocalConn)
            If myLocalConn.State = ConnectionState.Closed Then
                myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            myLocalConn.Close()
        End Using

    End Sub

    Private Shared Sub SyncCity()

        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, StateID, REPLACE(CityName, '''', '') AS CityName, ZIPCode FROM tblCity"
        Dim Query2 As String = "DELETE FROM tblCity; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        Dim rCount As Integer = dt.Rows.Count
        If rCount = 0 Then
            Exit Sub
        End If

        For y As Integer = 0 To rCount - 1 Step 1000
            Query2 &= "INSERT INTO tblCity (ID, StateID, CityName, ZIPCode) VALUES "
            For x As Integer = 0 To 999
                If x + y >= rCount Then
                    Exit For
                End If
                Query2 &= "(" & dt.Rows(x + y)(0).ToString & ", " & dt.Rows(x + y)(1).ToString & ", '" & dt.Rows(x + y)(2).ToString & "', '" & dt.Rows(x + y)(3).ToString & "'), "

            Next
            Query2 = Query2.Substring(0, Query2.Length - 2)

            Using cmd = New SqlCommand(Query2, myLocalConn)
                If myLocalConn.State = ConnectionState.Closed Then
                    myLocalConn.Open()
                End If
                cmd.ExecuteNonQuery()
                myLocalConn.Close()
            End Using
            
            Query2 = ""
        Next y

    End Sub
    Private Shared Sub SyncPonints()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, REPLACE(PointName, '''', '') AS PointName FROM tblESAPoint"
        Dim Query2 As String = "DELETE FROM tblESAPoint; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        Query2 &= "INSERT INTO tblESAPoint (ID, PointName) VALUES "
        For x As Integer = 0 To dt.Rows.Count - 1
            Query2 &= "(" & dt.Rows(x)(0).ToString & ", '" & dt.Rows(x)(1).ToString & "'), "
        Next
        Query2 = Query2.Substring(0, Query2.Length - 2)
        Query2 &= ";"

        Using cmd = New SqlCommand(Query2, myLocalConn)
            If myLocalConn.State = ConnectionState.Closed Then
                myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            myLocalConn.Close()
        End Using
        

    End Sub
    Private Shared Sub SyncCourses()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, REPLACE(CourseName, '''', '') AS CourseName, CourseCategory FROM tblCourses"
        Dim Query2 As String = "DELETE FROM tblCourses; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        Query2 &= "INSERT INTO tblCourses (ID, CourseName, CourseCategory) VALUES "
        For x As Integer = 0 To dt.Rows.Count - 1
            Query2 &= "(" & dt.Rows(x)(0).ToString & ", '" & dt.Rows(x)(1).ToString & "', " & dt.Rows(x)(2).ToString & "), "
        Next
        Query2 = Query2.Substring(0, Query2.Length - 2)
        Query2 &= ";"

        Using cmd = New SqlCommand(Query2, myLocalConn)
            If myLocalConn.State = ConnectionState.Closed Then
                myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            myLocalConn.Close()
        End Using
        

    End Sub
    Private Shared Sub SyncInstructor()
        Dim dt As New DataTable
        Dim Query1 As String = "SELECT ID, REPLACE(Name, '''', '') AS Name FROM tblInstructors"
        Dim Query2 As String = "DELETE FROM tblInstructors; "

        Try
            Using cmd = New SqlCommand(Query1, myOnlineConn)
                If myOnlineConn.State = ConnectionState.Closed Then
                    myOnlineConn.Open()
                End If
                Using dr As SqlDataReader = cmd.ExecuteReader
                    Try
                        dt.Load(dr)
                    Catch ex As Exception

                    End Try
                End Using
                myOnlineConn.Close()
            End Using
        Catch ex As Exception

        End Try

        If dt.Rows.Count = 0 Then
            Exit Sub
        End If
        Query2 &= "INSERT INTO tblInstructors (ID, Name) VALUES "
        For x As Integer = 0 To dt.Rows.Count - 1
            Query2 &= "(" & dt.Rows(x)(0).ToString & ", '" & dt.Rows(x)(1).ToString & "'), "
        Next
        Query2 = Query2.Substring(0, Query2.Length - 2)
        Query2 &= ";"

        Using cmd = New SqlCommand(Query2, myLocalConn)
            If myLocalConn.State = ConnectionState.Closed Then
                myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            myLocalConn.Close()
        End Using
        
    End Sub

    Public Shared Sub Synchronize()
        ShrinkDB("LocalDB")
        SyncCourseCategories()
        SyncCoursePlaces()
        SyncPonints()
        SyncCourses()
        SyncInstructor()
        SyncInstructorCourses()
        SyncCountry()
        SyncState()
        SyncCity()
    End Sub
End Class
