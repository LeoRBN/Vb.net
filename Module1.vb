Imports System.Data.OleDb
Imports System.Data

Module Module1
    Public conn As New OleDbConnection
    Public constr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\1. DATABASE\DMS1.accdb"
    Public cmd As New OleDbCommand
    Public dr As OleDbDataReader
    Public da As New OleDbDataAdapter
    Public ds As New DataSet()
    Public qry As String = Nothing
    Public maintable As DataGridView = MainForm.dgv1

    Sub connectDB()
        conn.Close()
        conn.ConnectionString = constr
        conn.Open()
    End Sub

    Public Function Load_TB_from_DB(qry As String, xfunction As String, tblename As String) As Boolean
        connectDB()

        If xfunction = Nothing Then
            ds.Tables.Clear()

            cmd = New OleDbCommand(qry, conn)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds, tblename)
            maintable.DataSource = ds.Tables(tblename)
#Region "LOGIN"
        ElseIf xfunction = "login" Then

            Dim idparam As New OleDbParameter("@userID", login.usertxt.Text)
            Dim pwordparam As New OleDbParameter("@pword", login.pwordtxt.Text)

            qry = "SELECT * FROM UserAcc WHERE UserID = @userID AND Pass_Word = @pword"
            cmd = New OleDbCommand(qry, conn)
            cmd.Parameters.Add(idparam)
            cmd.Parameters.Add(pwordparam)
            dr = cmd.ExecuteReader()

            dr.Read()


            If dr.HasRows Then
                Dim userName As String = dr("User_Name").ToString
                MainForm.Show()
                MessageBox.Show("Wellcome   " + userName)

            Else
                MessageBox.Show("User Not Found")
                login.pwordtxt.Clear()
                login.usertxt.Focus()


                Return False

            End If
#End Region

#Region "Register User"
        ElseIf xfunction = "register" Then

            Dim id As New OleDbParameter("@id", registrar.useridtxt.Text)
            Dim name As New OleDbParameter("@name", registrar.nametxt.Text)
            Dim pass As New OleDbParameter("@pass", registrar.pword2.Text)
            Dim lvl As New OleDbParameter("@lvl", registrar.lvl.Text)

            qry = "INSERT INTO UserAcc (UserID, User_Name, Pass_Word, UserLvl)"
            qry = qry + "VALUES (@id, @name, @pass, @lvl);"
            cmd = New OleDbCommand(qry, conn)
            cmd.Parameters.Add(id)
            cmd.Parameters.Add(name)
            cmd.Parameters.Add(pass)
            cmd.Parameters.Add(lvl)

            cmd.ExecuteNonQuery()
            MessageBox.Show("User Registered Succesfully!")
#End Region

#Region "UserList Control"
        ElseIf xfunction = "userlist" Then

            cmd = New OleDbCommand(qry, conn)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds, tblename)
            Userlist.dgvuserlist.DataSource = ds.Tables(tblename)

        End If
        Return True
    End Function
#End Region

#Region "Delete user"
    Public Function Delete_User(qry As String, xfunction As String, tblname As String) As Boolean

        connectDB()


        Return True
    End Function
#End Region

End Module
