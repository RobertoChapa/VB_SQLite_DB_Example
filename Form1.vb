'Database Project using SQLite
'
'4/15/2019
'Edited by Instructor 4/17/2019

Imports System.Data.SQLite
Imports System.Text.RegularExpressions

Public Class Form1

    ' Database Functions------------------------------------------------'

#Region "SelectData()"
    'Gets data from Database and fills DataGrid
    Private Function SelectData(sc As SQLiteConnection)

        Try
            sc.Open()

            Dim s = "select ID, Name, City, Telephone, Email, Gender from DatabaseTable order by ID desc"

            Dim cmdDataGrid As SQLiteCommand = New SQLiteCommand(s, sc)

            Dim da As New SQLiteDataAdapter With {
                .SelectCommand = cmdDataGrid
            }

            Dim dt As New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt

            Dim readerDataGrid As SQLiteDataReader = cmdDataGrid.ExecuteReader()
            readerDataGrid.Close()
            sc.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try

    End Function
#End Region

#Region "InsertData()"
    Private Function InsertData(sc As SQLiteConnection)

        Try

            If (TextBox2.Text.Length > 0) Then

                Dim s As String
                s = "insert into DatabaseTable (Name,City,Telephone,Email,Gender)"
                s += " Values "
                s += "(@Name, @City, @Telephone, @Email, @Gender)"

                Dim em, te As Boolean

                em = GetRegExpEmail(TextBox5.Text)
                te = GetRegExpPhoneNum(TextBox4.Text)

                If (em = False And String.IsNullOrEmpty(TextBox5.Text) = False) Then
                    MessageBox.Show("Email in wrong format")
                End If
                If (te = False And String.IsNullOrEmpty(TextBox4.Text) = False) Then
                    MessageBox.Show("Telephone number in wrong format")
                End If

                If ((em = True Or TextBox5.Text.Length < 1) And (te = True Or TextBox4.Text.Length < 1)) Then

                    Using cmd As New SQLiteCommand(s, sc)
                        cmd.Parameters.AddWithValue("@Name", TextBox2.Text)
                        cmd.Parameters.AddWithValue("@City", TextBox3.Text)
                        cmd.Parameters.AddWithValue("@Telephone", TextBox4.Text)
                        cmd.Parameters.AddWithValue("@Email", TextBox5.Text)
                        cmd.Parameters.AddWithValue("@Gender", ComboBox1.Text)

                        sc.Open()
                        cmd.ExecuteNonQuery()
                        sc.Close()

                    End Using

                Else

                    MessageBox.Show("Insert was NOT executed.")

                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Return 0

    End Function
#End Region

#Region "UpdateData()"
    Private Function UpdateData(sc As SQLiteConnection)

        Dim em, te, emr, ter As Boolean

        Dim s As String = ""

        s = "update DatabaseTable "
        s += " set "
        s += " Name = @Name, "
        s += " City = @City, "
        s += " Telephone = @Telephone, "
        s += " Email = @Email, "
        s += " Gender = @Gender "
        s += " where @ID = ID "

        Try

            For Each row As DataGridViewRow In DataGridView1.Rows

#Region "Email and Phone Formatting Checks"

                em = False
                emr = False
                te = False
                ter = False

                If (TryCast(row.Cells(4).Value, String) Is Nothing) Then
                    em = True
                End If

                If (TryCast(row.Cells(3).Value, String) Is Nothing) Then
                    te = True
                End If

                If em = False Then
                    emr = GetRegExpEmail(row.Cells(4).Value.ToString())
                End If

                If te = False Then
                    ter = GetRegExpPhoneNum(row.Cells(3).Value.ToString())
                End If

                If (emr = False And TryCast(row.Cells(4).Value, String) IsNot Nothing) Then
                    MessageBox.Show("Email is wrong format")
                End If

                If (ter = False And TryCast(row.Cells(3).Value, String) IsNot Nothing) Then
                    MessageBox.Show("Telephone number is wrong format")
                End If
#End Region

                If ((emr = True Or TryCast(row.Cells(4).Value, String) Is Nothing) And (ter = True Or TryCast(row.Cells(3).Value, String) Is Nothing)) Then

                    Using cmd As New SQLiteCommand(s, sc)
                        cmd.Parameters.AddWithValue("@ID", row.Cells(0).Value)
                        cmd.Parameters.AddWithValue("@Name", row.Cells(1).Value)
                        cmd.Parameters.AddWithValue("@City", row.Cells(2).Value)
                        cmd.Parameters.AddWithValue("@Telephone", row.Cells(3).Value)
                        cmd.Parameters.AddWithValue("@Email", row.Cells(4).Value)
                        cmd.Parameters.AddWithValue("@Gender", row.Cells(5).Value)

                        sc.Open()
                        cmd.ExecuteNonQuery()
                        sc.Close()

                    End Using
                Else

                    MessageBox.Show("Update was NOT executed.")
                    Exit For

                End If

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function
#End Region

#Region "DeleteData()"
    Private Function DeleteData(sc As SQLiteConnection)

        Dim s As String = "delete from DatabaseTable where @ID = ID"

        If DataGridView1.SelectedRows.Count > 0 Then

            Using cmd As New SQLiteCommand(s, sc)

                cmd.Parameters.AddWithValue("@ID", DataGridView1.SelectedRows(0).Cells(0).Value)

                sc.Open()
                cmd.ExecuteNonQuery()
                sc.Close()

            End Using

        End If

        Return 0

    End Function
#End Region

    ' Auxillary-----------------------------------------------------------'

#Region "GetCS()"
    'Gets Connection String for SQLite DB
    'The database must be in the directory ..\Database Project\SQLite Database\DatabaseTable.db
    Private Function GetCS() As SQLiteConnection

        Dim cs = System.IO.Directory.GetCurrentDirectory()
        cs = cs.Replace("\Database Project\bin\Debug", "\SQLite Database\DatabaseTable.db")
        cs = "Data Source = " + cs

        Dim SQLConn As New SQLiteConnection(cs)
        SQLConn = New SQLiteConnection(cs)

        Return SQLConn

    End Function
#End Region

#Region "Validation"
    Function GetRegExpEmail(s As String) As Boolean

        Dim t As Boolean = False

        If (Regex.IsMatch(s, "^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")) Then
            t = True
        Else
            t = False
        End If

        Return t

    End Function

    Function GetRegExpPhoneNum(s As String) As Boolean

        Dim t As Boolean = False

        If (Regex.IsMatch(s, "((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4}")) Then
            t = True
        Else
            t = False
        End If

        Return t

    End Function
#End Region

#Region "clearTextBoxes()"
    Function clearTextBoxes()

        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        ComboBox1.SelectedIndex = -1
        TextBox2.Select()

        Return 0
    End Function
#End Region

    'Event Handlers--------------------------------------------------------'

#Region "Event Handlers"
    ' Insert Button
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, MyBase.Load
        Dim sc As SQLiteConnection = GetCS() ' Added by Instructor 
        InsertData(sc)
        clearTextBoxes()
        SelectData(sc)
    End Sub

    ' Update 
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim sc As SQLiteConnection = GetCS()
        UpdateData(sc)
    End Sub

    ' Delete
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sc As SQLiteConnection = GetCS()
        DeleteData(sc)
        SelectData(sc)
    End Sub

    ' Close Button
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
#End Region

End Class


