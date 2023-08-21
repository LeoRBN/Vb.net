Imports System.Data.OleDb

Public Class MainForm
    Dim conn As New OleDbConnection
    Dim qry As String = Nothing
    Dim Cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As New OleDbDataAdapter
    Dim refTB As New DataTable()
    Dim str As String = Nothing

#Region "DATABASE CONNECTION"
    Sub ConnectDB()
        Try
            With conn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\1. DATABASE\DMS1.accdb;Persist Security Info=False;"
                .Open()

            End With
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "DATABASE ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()

        End Try
    End Sub
#End Region

#Region "DISPLAY DATABASE ON LISTVIEW"

    Sub loadDB()

        If str = "add" Or str = "edit" Or str = "clear" Or str = Nothing Then

            Lv1.Items.Clear()
            qry = "SELECT * FROM ST_people"
            Cmd = New OleDbCommand(qry, conn)
            dr = Cmd.ExecuteReader
            While dr.Read
                With Lv1
                    .Items.Add(dr("ID"))
                    With .Items(.Items.Count - 1).SubItems
                        .Add(dr("STM_ID"))
                        .Add(dr("name"))
                        .Add(dr("position"))
                        .Add(dr("area"))
                    End With
                End With
            End While

        ElseIf str = "search" Then

            Lv1.Items.Clear()
            Try
                qry = "SELECT * FROM ST_people WHERE ( 1 = 1) "

                If Not String.IsNullOrEmpty(fcb2.Text) Then
                    qry &= " AND area = '" & fcb2.Text & "'"
                End If

                If Not String.IsNullOrEmpty(fcb1.Text) Then
                    qry &= " AND [position] = '" & fcb1.Text & "'"
                End If

                If Not String.IsNullOrEmpty(ftb1.Text) Then
                    qry &= " AND STM_id = " & ftb1.Text & ""
                End If

                Cmd = New OleDbCommand(qry, conn)
                dr = Cmd.ExecuteReader

                While dr.Read
                    With Lv1
                        .Items.Add(dr("ID"))
                        With .Items(.Items.Count - 1).SubItems
                            .Add(dr("STM_ID"))
                            .Add(dr("name"))
                            .Add(dr("position"))
                            .Add(dr("area"))
                        End With
                    End With

                End While
                MessageBox.Show("Filtered")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub

    Sub lv1set()
        With Lv1.Columns
            .Add("ID", 0, HorizontalAlignment.Left)
            .Add("STM ID", 70, HorizontalAlignment.Left)
            .Add("User Name", 150, HorizontalAlignment.Left)
            .Add("Position", 100, HorizontalAlignment.Left)
            .Add("Area", 100, HorizontalAlignment.Left)
        End With
    End Sub

#End Region

#Region "CONTROL SETTING"
    Sub set1()
        Dim disAB() As Control = {tb1, tb2, cb1, cb2, savebtn, deletebtn, editbtn}
        For Each control As Control In disAB
            control.Enabled = False
        Next

    End Sub
    Sub set2()
        Dim enab() As Control = {tb1, tb2, cb1, cb2, savebtn, deletebtn}
        For Each control As Control In enab
            control.Enabled = True
        Next

    End Sub

    Sub comboX()

        If str = "area" Then

            fcb2.Items.Clear()
            cb2.Items.Clear()

            qry = "SELECT area FROM Area"
            Cmd = New OleDbCommand(qry, conn)
            dr = Cmd.ExecuteReader
            While dr.Read
                cb2.Items.Add(dr("area"))
                fcb2.Items.Add(dr("area"))
            End While

        ElseIf str = "position" Then

            cb1.Items.Clear()
            fcb1.Items.Clear()

            qry = "SELECT [position] FROM UserSecurityLvl"
            Cmd = New OleDbCommand(qry, conn)
            dr = Cmd.ExecuteReader
            While dr.Read
                cb1.Items.Add(dr("position"))
                fcb1.Items.Add(dr("position"))
            End While

        End If

    End Sub
#End Region

#Region "CLEAR ALL TEXT"
    Sub clearol()
        Dim cl() As Control = {tb1, tb2, ftb1}
        For Each control As Control In cl
            control.Text = ""
        Next
        cb1.SelectedIndex = -1
        cb2.SelectedIndex = -1
        fcb1.SelectedIndex = -1
        fcb2.SelectedIndex = -1

    End Sub


#End Region

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ConnectDB()
        loadDB()
        lv1set()
        comboX()
        set1()

    End Sub

    Private Sub Lv1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lv1.SelectedIndexChanged
        editbtn.Enabled = True
        deletebtn.Enabled = True
        If Lv1.SelectedItems.Count > 0 Then
            With Lv1.SelectedItems(0)
                tb1.Text = .SubItems(1).Text
                tb2.Text = .SubItems(2).Text
                cb1.Text = .SubItems(3).Text
                cb2.Text = .SubItems(4).Text
                uniqID.Text = .SubItems(0).Text
            End With
        End If
    End Sub

    Private Sub editbtn_Click(sender As Object, e As EventArgs) Handles editbtn.Click
        str = "edit"

        If Lv1.SelectedItems.Count > 0 Then
            set2()
            addbtn.Enabled = False
            deletebtn.Enabled = False
        End If


    End Sub

    Private Sub addbtn_Click(sender As Object, e As EventArgs) Handles addbtn.Click
        str = "add"
        set2()
        clearol()
        editbtn.Enabled = False
        deletebtn.Enabled = False

    End Sub

    Private Sub clearbtn_Click(sender As Object, e As EventArgs) Handles clearbtn.Click
        str = "clear"
        clearol()
        set1()
        loadDB()
        addbtn.Enabled = True

    End Sub

    Private Sub savebtn_Click(sender As Object, e As EventArgs) Handles savebtn.Click

        If str = "add" Then

            qry = "INSERT INTO ST_people (STM_id, name, [position], area)"
            qry = qry + " VALUES ('" & tb1.Text & "','" & tb2.Text & "','" & cb1.Text & "','" & cb2.Text & "')"
            Cmd = New OleDbCommand
            Try
                With Cmd
                    .CommandText = qry
                    .Connection = conn
                    .ExecuteNonQuery()
                End With
                MessageBox.Show("Succesfully Added")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        ElseIf str = "edit" Then

            qry = "UPDATE ST_people SET"
            qry = qry + " STM_id = '" & tb1.Text & "',"
            qry = qry + " name = '" & tb2.Text & "',"
            qry = qry + " [position] = '" & cb1.Text & "',"
            qry = qry + " area = '" & cb2.Text & "'"
            qry = qry + " WHERE "
            qry = qry + " ID = " & uniqID.Text
            Try
                Cmd = New OleDbCommand(qry, conn)
                Cmd.ExecuteNonQuery()
                MessageBox.Show("Update Succesfully!")

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try



        End If
        loadDB()
        clearol()
        set1()

    End Sub

    Private Sub deletebtn_Click(sender As Object, e As EventArgs) Handles deletebtn.Click
        If MsgBox("Are you sure do you want to delete this record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            qry = "DELETE FROM ST_people WHERE ID = " & uniqID.Text
            Cmd = New OleDbCommand(qry, conn)
            Cmd.ExecuteNonQuery()
            MessageBox.Show("RECORD DELETED")
        End If
        deletebtn.Enabled = False
        clearol()
        loadDB()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles srchbtn.Click
        str = "search"
        loadDB()

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        str = "clear"
        clearol()
        loadDB()

    End Sub

    Private Sub cb1_Click(sender As Object, e As EventArgs) Handles cb1.Click, fcb1.Click
        str = "position"
        comboX()

    End Sub

    Private Sub cb2_Click(sender As Object, e As EventArgs) Handles cb2.Click, fcb2.Click
        str = "area"
        comboX()

    End Sub
End Class
