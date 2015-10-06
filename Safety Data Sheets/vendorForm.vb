Public Class vendorForm

    Private Sub vendorForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        vendorNameTextBox.Text = ""
    End Sub

    Private Sub saveSettingsButton_Click(sender As Object, e As EventArgs) Handles saveSettingsButton.Click

        'Check if required information is blank
        If vendorNameTextBox.Text = "" Then
            MsgBox("Please enter a Vendor.")
            Exit Sub
        End If

        'Connect to the database
        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter
        Dim Sql As String

        'Provider to access the database and where the database is located
        dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = " & My.Settings.SDSDBLocation

        con.ConnectionString = dbProvider & dbSource

        'Opening the database connection
        con.Open()

        'Search string to select all Tags from the tagTable
        Sql = "SELECT * FROM VendorList"

        'Send the search to the data adapter
        da = New OleDb.OleDbDataAdapter(Sql, con)

        'Tell the data adapter to fill the dataset
        da.Fill(ds, "VendorList")

        'Add a new row to the dataset
        'Create a command builder 
        Dim cb As New OleDb.OleDbCommandBuilder(da)
        Dim dsNewRow As DataRow

        'Create new row in the dataset
        dsNewRow = ds.Tables("VendorList").NewRow()

        'Add the new tag to the new datarow in the dataset
        dsNewRow.Item("VendorName") = vendorNameTextBox.Text

        'Add the row to the dataset
        ds.Tables("VendorList").Rows.Add(dsNewRow)

        'Update the database using the data adapter
        da.Update(ds, "VendorList")

        'MsgBox("Database is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database is now Closed")

        vendorNameTextBox.Text = ""
        Me.Close()

    End Sub

End Class