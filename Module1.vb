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
    Public logIN_p As String = Nothing


    Sub connectDB()
        conn.Close()
        conn.ConnectionString = constr
        conn.Open()
    End Sub

    Sub xdelete()

        Dim selectedRow As DataGridViewRow = Userlist.dgvuserlist.SelectedRows(0)
        qry = "DELETE FROM UserAcc WHERE UserID = @primarykey"
        cmd = New OleDbCommand(qry, conn)
        cmd.Parameters.AddWithValue("@primarykey", selectedRow.Cells("UserID").Value)
        cmd.ExecuteNonQuery()
        Userlist.dgvuserlist.Rows.Remove(selectedRow)

    End Sub

#Region "LOAD_TB_From_DB"

    Public Function Load_TB_from_DB(qry As String, xfunction As String, tblename As String) As Boolean
        connectDB()
        ds.Tables.Clear()
        MainForm.tableName.Text = tblename
        If xfunction = Nothing Then

            cmd = New OleDbCommand(qry, conn)
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds, tblename)
            maintable.DataSource = ds.Tables(tblename)

#Region "LOGIN"
        ElseIf xfunction = "login" Then

            Dim idparam As New OleDbParameter("@userID", login.usertxt.Text)
            Dim pwordparam As New OleDbParameter("@pword", login.pwordtxt.Text)
            Dim idparam1 As New OleDbParameter("@userID", login_2.userID.Text)
            Dim pwordparam1 As New OleDbParameter("@pword", login_2.pword.Text)
            Dim ds1 As New DataSet
            qry = "SELECT * FROM UserAcc WHERE UserID = @userID AND Pass_Word = @pword"
            cmd = New OleDbCommand(qry, conn)

            If logIN_p = "login1" Then
                cmd.Parameters.Add(idparam)
                cmd.Parameters.Add(pwordparam)
            ElseIf logIN_p = "login2" Then
                cmd.Parameters.Add(idparam1)
                cmd.Parameters.Add(pwordparam1)
            End If
            da = New OleDbDataAdapter(cmd)
            da.Fill(ds1, "UserID")
            dr = cmd.ExecuteReader()
            dr.Read()


            If dr.HasRows Then
                Dim userName As String = dr("User_Name").ToString

                If logIN_p = "login1" Then
                    MainForm.Show()

                    Dim row As DataRow = ds1.Tables("UserID").Rows(0)
                    MainForm.UserMenu.Text = row("User_Name").ToString()

                    MessageBox.Show("Wellcome   " + userName)
                ElseIf logIN_p = "login2" Then
                    login_2.Close()
                    xdelete()
                    MessageBox.Show("Deleted Succesfully")
                End If

            Else
                MessageBox.Show("User Not Found")
                login.pwordtxt.Clear()
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

#Region "UPDATE"
        ElseIf xfunction = "Update" Then

            Dim fldname As String = Nothing
            Dim selectedROw As DataGridViewRow = MainForm.dgv1.SelectedRows(0)

            Select Case tblename
                Case tblename = "Area"
                    fldname = "ID"
                Case tblename = "ST_people"
                    fldname = "STM_id"
                Case Else : Exit Select
            End Select

            Dim item As String = selectedROw.Cells(fldname).Value.ToString
            qry = "SELECT * FROM " + tblename
            qry = "WHERE " + fldname + "=" + item
            cmd = New OleDbCommand(qry, conn)
            cmd.ExecuteNonQuery()

        End If
#End Region
        Return True


    End Function
#End Region

#End Region

#Region "Delete user"
    Public Function Delete_User(qry As String, xfunction As String, tblname As String) As Boolean




        Return True
    End Function
#End Region

        

'THIS CODE WILL UPLOAD A CSV FILE TO THE DATABASE. MAKE USE THE CULOMN HEADER OF CSV IS SAME WITH THE DATABASE

        
'CODE FOR FORM

        Private Sub selectfile_Click(sender As Object, e As EventArgs) Handles selectfile.Click

    Dim openfiledialog As New OpenFileDialog
    openfiledialog.Filter = "CSV files (*.csv)|.csv"

    If openfiledialog.ShowDialog() = DialogResult.OK Then
        Dim selectedfile As String = openfiledialog.FileName

        Using reader As New StreamReader(selectedfile)
            Dim dt As New DataTable()
            Dim isfirstrow As Boolean = True

            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                Dim header() As String = line.Split(",")

                If dt.Columns.Count = 0 Then
                    For Each head As String In header
                        dt.Columns.Add(head)
                    Next
                Else
                    dt.Rows.Add(header)

                End If
            End While
            dgvupload.DataSource = dt
            rcount.Text = dgvupload.Rows.Count

        End Using
    End If

End Sub

'CODE FOR MODULE

         Public Function UPLOAD_DATA(qry As String, xfunction As String, esdtype As String, esdstatus As String) As Boolean

     connectDB()
     Dim esdUpload As DataGridView = Usercontrol.dgvupload
     Dim dt As New DataTable()

     For Each col As DataGridViewColumn In esdUpload.Columns
         dt.Columns.Add(col.HeaderText, col.ValueType)

     Next
     For Each row As DataGridViewRow In esdUpload.Rows
         Dim rowDT As DataRow = dt.Rows.Add()
         For Each cell As DataGridViewCell In row.Cells
             rowDT(cell.ColumnIndex) = cell.Value

         Next
     Next


     qry = "SELECT * FROM ESD_WS"
     Using da As New OleDbDataAdapter()
         da.SelectCommand = New OleDbCommand(qry, conn)
         Using cmdbuilder As New OleDbCommandBuilder(da)
             da.Update(dt)

         End Using
         MessageBox.Show("UPLOAD success")
     End Using

     Return True
End Function


End Module

