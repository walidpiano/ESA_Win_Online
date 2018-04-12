Imports System.Data.SqlClient

Public Class Store


End Class

Public Class Mail
    Public ID As Int64
    Public Subject As String
    Public Type As String
    Public Sent As Boolean
    Public MailDate As Date
    Public Content As String

    Public Sub New(ByVal ID As Int64, ByVal Subject As String, ByVal Type As String, ByVal Sent As Boolean _
                   , ByVal MailDate As Date, ByVal Content As String)
        Me.ID = ID
        Me.Subject = Subject
        Me.Type = Type
        Me.Sent = Sent
        Me.MailDate = MailDate
        Me.Content = Content

    End Sub

    Public Sub Add()
        Dim Query As String = "INSERT INTO tblMails ([Subject], [Type], [Sent], [Date], Content) VALUES (@Subject, @Type, @Sent, @Date, @Content); SELECT @@IDENTITY;"
        Dim result As Int64 = 0
        Using cmd = New SqlCommand(Query, ExClass.myLocalConn)
            cmd.Parameters.AddWithValue("@Subject", Me.Subject)
            cmd.Parameters.AddWithValue("@Type", Me.Type)
            cmd.Parameters.AddWithValue("@Sent", Me.Sent)
            cmd.Parameters.AddWithValue("@Date", Me.MailDate)
            cmd.Parameters.AddWithValue("@Content", Me.Content)

            If ExClass.myLocalConn.State = ConnectionState.Closed Then
                ExClass.myLocalConn.Open()
            End If
            result = cmd.ExecuteScalar()
            ExClass.myLocalConn.Close()
        End Using

        Me.ID = result

    End Sub

    Public Shared Function Recall(ByVal ID As Int64) As String
        Dim result As String = ""
        Dim Query As String = "SELECT Content FROM tblMails WHERE ID = " & ID.ToString & ";"
        Using cmd = New SqlCommand(Query, ExClass.myLocalConn)
            If ExClass.myLocalConn.State = ConnectionState.Closed Then
                ExClass.myLocalConn.Open()
            End If
            Using dr As SqlDataReader = cmd.ExecuteReader
                If dr.Read() Then
                    result = dr(0)
                End If
            End Using
            ExClass.myLocalConn.Close()
        End Using

        Return result
    End Function

    Public Sub Update()
        Dim Query As String = "UPDATE tblMails SET [Subject] = @Subject, [Type] = @Type, [Sent] = @Sent" _
                              & ", Content = @Content " _
                              & " WHERE ID = " & Me.ID & ";"
        Using cmd = New SqlCommand(Query, ExClass.myLocalConn)
            cmd.Parameters.AddWithValue("@Subject", Me.Subject)
            cmd.Parameters.AddWithValue("@Type", Me.Type)
            cmd.Parameters.AddWithValue("@Sent", Me.Sent)
            cmd.Parameters.AddWithValue("@Date", Me.MailDate)
            cmd.Parameters.AddWithValue("@Content", Me.Content)

            If ExClass.myLocalConn.State = ConnectionState.Closed Then
                ExClass.myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            ExClass.myLocalConn.Close()
        End Using

    End Sub

    Public Shared Sub Delete(ByVal ID As Int64)
        Dim Query As String = "DELETE FROM tblMails WHERE ID = " & ID & ";"
        Using cmd = New SqlCommand(Query, ExClass.myLocalConn)
            If ExClass.myLocalConn.State = ConnectionState.Closed Then
                ExClass.myLocalConn.Open()
            End If
            cmd.ExecuteNonQuery()
            ExClass.myLocalConn.Close()
        End Using
    End Sub

    Public Shared Function getAllMails() As DataTable
        Dim Query As String = "SELECT ID, [Subject], Sent, [Date] FROM tblMails;"
        Dim dt As New DataTable
        dt.Columns.Add("ID", GetType(Long))
        dt.Columns.Add("Subject", GetType(String))
        dt.Columns.Add("Sent", GetType(Short))
        dt.Columns.Add("Date", GetType(Date))

        Using cmd = New SqlCommand(Query, ExClass.myLocalConn)
            ExClass.toConnect(True, True)
            Using dr As SqlDataReader = cmd.ExecuteReader
                dt.Load(dr)
            End Using
            ExClass.myLocalConn.Close()
        End Using
        Return dt
    End Function
End Class
