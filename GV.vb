Public Class GV
    Public Shared myLocalConn As String = "Data Source=(LocalDB)\v11.0;AttachDbFilename=" & My.Settings.DatabasePath & ";Integrated Security=True"

    Public Shared myOnlineConn As String = "workstation id=ESAOnlineDB.mssql.somee.com;packet size=4096;user id=walid2_SQLLogin_1;pwd=xpe82xp4a2;data source=ESAOnlineDB.mssql.somee.com;persist security info=False;initial catalog=ESAOnlineDB"

End Class
