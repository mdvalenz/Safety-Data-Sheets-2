Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Access

Public Class mainForm

    'Dim materialID As Integer = Nothing
    Dim materialName As String = Nothing
    Dim newMaterial As Boolean = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'SDS_LogDataSet1.SDS_Log' table. You can move, or remove it, as needed.
        Me.SDS_LogTableAdapter.Fill(Me.SDS_LogDataSet1.SDS_Log)
        'TODO: This line of code loads data into the 'SDS_LogDataSet.SDS_List' table. You can move, or remove it, as needed.
        Me.SDS_ListTableAdapter.Fill(Me.SDS_LogDataSet.SDS_List)

        'Fill all datagridviews and dropdown lists
        Call loadDefaultData()
        My.Settings.materialName = ""

    End Sub

    Private Sub loadDefaultData()

        'Fill DataGridViews with the SDS List and SDS Log
        Call loadSDSLog()

        'Fill the Material List dropdown from the database
        Call loadMaterials()

    End Sub

    Private Sub SettingsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem1.Click
        settingsForm.Show()
    End Sub

    Private Sub loadMaterials()

        'Connect to the task database
        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String
        Dim ds As New DataSet
        Dim Sql As String

        'Provider to access the database and where the database is located
        dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = " & My.Settings.SDSDBLocation

        con.ConnectionString = dbProvider & dbSource

        'Opening the database connection
        con.Open()

        'Search string to select all tags from the tagTable
        Sql = "SELECT ID, Material FROM SDS_List ORDER by Material"
        Dim cm As New OleDb.OleDbCommand(Sql, con)

        'Send the search to the data reader
        Dim dr As OleDb.OleDbDataReader = cm.ExecuteReader

        'Get the database into the dropdown list
        materialSDSLookupComboBox.Items.Clear()
        SDSLogComboBox.Items.Clear()

        While dr.Read
            materialSDSLookupComboBox.Items.Add(dr("Material"))
            SDSLogComboBox.Items.Add(dr("Material"))
        End While

        'Add instructions to the dropdown list and set as default selection
        materialSDSLookupComboBox.Items.Insert(0, "Enter Material")
        materialSDSLookupComboBox.SelectedIndex = 0

        SDSLogComboBox.Items.Insert(0, "Enter Material")
        SDSLogComboBox.SelectedIndex = 0

        'Closing the database connection
        dr.Close()
        con.Close()

        'MsgBox("Database is now Closed")

    End Sub

    Private Sub loadSDSLog()

        'Connect to the database
        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim bs As New BindingSource
        Dim da As OleDb.OleDbDataAdapter
        Dim Sql As String

        'Provider to access the database and where the database is located
        dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = " & My.Settings.SDSDBLocation

        con.ConnectionString = dbProvider & dbSource

        'Opening the database connection
        con.Open()

        Sql = "SELECT * FROM SDS_Log ORDER BY DateLogged DESC"

        'Send the search to the data adapter
        da = New OleDb.OleDbDataAdapter(Sql, con)

        'Tell the data adapter to fill the dataset
        da.Fill(dt)
        bs.DataSource = dt
        DataGridView4.DataSource = bs
        DataGridView4.Columns("ID").Visible = False

        DataGridView4.Columns(1).HeaderCell.Value = "Material"
        DataGridView4.Columns(2).HeaderCell.Value = "Date Added/Revised"

        DataGridView4.Columns(2).DisplayIndex = 1
        DataGridView4.Columns(1).DisplayIndex = 2

        da.Update(dt)

        'MsgBox("Database is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database is now Closed")

    End Sub

    Private Sub DataGridView4_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellClick

        'Get Task ID number
        If e.RowIndex <> -1 Then
            My.Settings.materialName = DataGridView4.Rows(e.RowIndex).Cells(1).Value
        End If

    End Sub

    Private Sub DataGridView4_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.CellMouseDoubleClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            My.Settings.materialName = DataGridView4.Rows(e.RowIndex).Cells(1).Value
            Call getMaterial()
            mainTabControl.SelectTab(0)
        End If
    End Sub

    Private Sub searchSDSLog()
        'Get search string
        Dim searchString = SDSLogComboBox.Text

        'Connect to the database
        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim bs As New BindingSource
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
        Dim dc As OleDb.OleDbCommand
        Dim Sql As String
        Dim buildString As String = ""

        'Provider to access the database and where the database is located
        dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = " & My.Settings.SDSDBLocation

        con.ConnectionString = dbProvider & dbSource

        'Opening the database connection
        con.Open()

        'Search string
        'only show tags that match search, all tags if blank
        If searchString = "" Then
            Sql = "SELECT * FROM SDS_Log ORDER by DateLogged"
        Else
            Sql = "SELECT * FROM SDS_Log WHERE Material LIKE @searchString ORDER BY DateLogged DESC"
        End If

        'Send the search to the data adapter
        dc = New OleDb.OleDbCommand(Sql, con)

        dc.Parameters.AddWithValue("@searchString", "%" & searchString & "%")
        da.SelectCommand = dc

        'Tell the data adapter to fill the dataset
        da.Fill(dt)
        bs.DataSource = dt
        DataGridView4.DataSource = bs

        da.Update(dt)

        'MsgBox("Database is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database is now Closed")
    End Sub

    Private Sub SDSLogComboBox_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles SDSLogComboBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call searchSDSLog()
        End If
    End Sub

    Private Sub refreshSDSLogButton_Click(sender As Object, e As EventArgs) Handles refreshSDSLogButton.Click
        SDSLogComboBox.Text = ""
        Call loadSDSLog()
    End Sub

    Private Sub searchSDSLogButton_Click(sender As Object, e As EventArgs) Handles searchSDSLogButton.Click
        Call searchSDSLog()
    End Sub

    Private Sub getInfoSDSLogButton_Click(sender As Object, e As EventArgs) Handles getInfoSDSLogButton.Click
        Call getMaterial()
        mainTabControl.SelectTab(0)
    End Sub

    Private Sub getMaterial()
        materialSDSLookupComboBox.Text = My.Settings.materialName
        Call getMaterialInformation()
    End Sub

    Private Sub materialSDSLookupComboBox_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles materialSDSLookupComboBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call getMaterialInformation()
        End If
    End Sub

    Private Sub getMaterialInformation()

        'Check if anything was entered
        If materialSDSLookupComboBox.Text = "Enter Material" Then
            MsgBox("Please enter a material", MsgBoxStyle.Exclamation, "Enter Material")
            Exit Sub
        End If

        'Get Material Name
        materialName = materialSDSLookupComboBox.Text

        'Search for material in SDS list
        'Connect to the task database
        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String
        Dim ds As New DataSet
        Dim Sql As String

        'Provider to access the database and where the database is located
        dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = " & My.Settings.SDSDBLocation

        con.ConnectionString = dbProvider & dbSource

        'Opening the database connection
        con.Open()

        'Search string to select all tags from the tagTable
        Sql = "SELECT * FROM SDS_List WHERE Material=@materialName"

        Dim cm As New OleDb.OleDbCommand(Sql, con)
        cm.Parameters.AddWithValue("@materialName", materialName)

        'Send the search to the data reader
        Dim dr As OleDb.OleDbDataReader = cm.ExecuteReader

        Try
            If dr.Read() Then
                'Get the database into the dropdown list
                'While dr.Read
                'materialSDSLookupComboBox.Text = dr.Item(1).ToString
                CASNumberSDSLookupTextBox.Text = dr.Item(3).ToString
                hazardsSDSLookupTextBox.Text = dr.Item(2).ToString

                vendor1TextBox.Text = dr.Item(4).ToString
                If vendor1TextBox.Text <> "" Then vendor1EDSDSLookupDateTimePicker.Value = dr.Item(5)
                vendor2TextBox.Text = dr.Item(6).ToString
                If vendor2TextBox.Text <> "" Then vendor2EDSDSLookupDateTimePicker.Value = dr.Item(7)

                'End While

                'MsgBox("Database is now Closed")
                newMaterial = False
            Else
                '    Dim inputResult As Integer = MsgBox("Material Not Found. Would you like to add a new material?", MsgBoxStyle.YesNo, "Add New Material?")

                '    If inputResult = 6 Then

                'Clear the rest of the form
                CASNumberSDSLookupTextBox.Text = ""
                hazardsSDSLookupTextBox.Text = ""
                vendor1TextBox.Text = ""
                vendor1EDSDSLookupDateTimePicker.Value = Date.Today
                vendor2TextBox.Text = ""
                vendor2EDSDSLookupDateTimePicker.Value = Date.Today

                'MsgBox("Please enter the new material's information.", , "Enter New Material")
                newMaterial = True
                'Else
                'Call clearForm()
                'End If
            End If

        Finally

            'Closing the database connection
            dr.Close()
            con.Close()

        End Try

    End Sub

    Private Sub clearForm()

        materialSDSLookupComboBox.SelectedIndex = 0
        CASNumberSDSLookupTextBox.Text = ""
        hazardsSDSLookupTextBox.Text = ""
        vendor1TextBox.Text = ""
        vendor1EDSDSLookupDateTimePicker.Value = Date.Today
        vendor2TextBox.Text = ""
        vendor2EDSDSLookupDateTimePicker.Value = Date.Today

    End Sub

    Private Sub clearSDSLookupButton_Click(sender As Object, e As EventArgs) Handles clearSDSLookupButton.Click
        Call clearForm()
    End Sub

    Private Sub resetSDSLookupButton_Click(sender As Object, e As EventArgs) Handles resetSDSLookupButton.Click
        Call getMaterialInformation()
    End Sub

    Private Sub saveSDSLookupButton_Click(sender As Object, e As EventArgs) Handles saveSDSLookupButton.Click

        If newMaterial = False Then
            My.Settings.materialName = materialSDSLookupComboBox.Text
            Call editMaterial()
        Else
            Call saveMaterial()
        End If

        Call addlog()
        Call loadMaterials()
        Call loadSDSLog()

    End Sub

    Private Sub editMaterial()

        'Check if required information is blank
        If materialSDSLookupComboBox.Text = "" Or CASNumberSDSLookupTextBox.Text = "" Or hazardsSDSLookupTextBox.Text = "" Or vendor1TextBox.Text = "" Then
            MsgBox("Please enter all required information.")
            Exit Sub
        End If

        'Get Variables
        materialName = My.Settings.materialName
        Dim hazards As String = hazardsSDSLookupTextBox.Text
        Dim CASNumber As String = CASNumberSDSLookupTextBox.Text
        Dim vendor1 As String = vendor1TextBox.Text
        Dim vendor1Date As Date = Format(vendor1EDSDSLookupDateTimePicker.Value, "M/d/yyyy")
        Dim vendor2 As String = vendor2TextBox.Text
        Dim vendor2Date As Date = Format(vendor2EDSDSLookupDateTimePicker.Value, "M/d/yyyy")

        'Connect to the task database
        Dim con As New OleDb.OleDbConnection
        Dim dbProvider As String
        Dim dbSource As String
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
        Dim dc As OleDb.OleDbCommand
        Dim Sql As String

        'Provider to access the database and where the database is located
        dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
        dbSource = "Data Source = " & My.Settings.SDSDBLocation

        con.ConnectionString = dbProvider & dbSource

        'Opening the database connection
        con.Open()

        Try
            'Base SQL String
            Sql = "UPDATE SDS_List SET Hazards=@hazards, CASNumber=@CASNumber, "

            If vendor1 <> "" Then
                Sql = Sql & "Vendor1=@vendor1, ED1=@vendor1Date, "
            End If

            If vendor2 <> "" Then
                Sql = Sql & "Vendor2=@vendor2, ED2=@vendor2Date, "
            End If

            Sql = Sql & "WHERE Material='" & materialName & "'"

            'Send the search to the data adapter
            dc = New OleDb.OleDbCommand(Sql, con)

            'dc.Parameters.AddWithValue("@Material", materialName)
            dc.Parameters.AddWithValue("@hazards", hazards)
            dc.Parameters.AddWithValue("@CASNumber", CASNumber)

            If vendor1 <> "" Then
                dc.Parameters.AddWithValue("@vendor1", vendor1)
                dc.Parameters.AddWithValue("@vendor1Date", vendor1Date)
            End If

            If vendor2 <> "" Then
                dc.Parameters.AddWithValue("@vendor2", vendor2)
                dc.Parameters.AddWithValue("@vendor2Date", vendor2Date)
            End If

            dc.ExecuteNonQuery()

        Catch ex As Exception

        End Try

        'Closing the database connection
        con.Close()

    End Sub

    Private Sub saveMaterial()

        'Check if required information is blank
        If materialSDSLookupComboBox.Text = "" Or CASNumberSDSLookupTextBox.Text = "" Or hazardsSDSLookupTextBox.Text = "" Or vendor1TextBox.Text = "" Then
            MsgBox("Please enter all required information.")
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
        Sql = "SELECT * FROM SDS_List"

        'Send the search to the data adapter
        da = New OleDb.OleDbDataAdapter(Sql, con)

        'Tell the data adapter to fill the dataset
        da.Fill(ds, "SDS")

        'Add a new row to the dataset
        'Create a command builder 
        Dim cb As New OleDb.OleDbCommandBuilder(da)
        Dim dsNewRow As DataRow

        'Create new row in the dataset
        dsNewRow = ds.Tables("SDS").NewRow()

        'Add the new tag to the new datarow in the dataset
        dsNewRow.Item("Material") = materialSDSLookupComboBox.Text
        dsNewRow.Item("CASNumber") = CASNumberSDSLookupTextBox.Text
        dsNewRow.Item("Hazards") = hazardsSDSLookupTextBox.Text

        If vendor1TextBox.Text <> "" Then
            dsNewRow.Item("Vendor1") = vendor1TextBox.Text
            dsNewRow.Item("ED1") = vendor1EDSDSLookupDateTimePicker.Value
        End If

        If vendor2TextBox.Text <> "" Then
            dsNewRow.Item("Vendor2") = vendor2TextBox.Text
            dsNewRow.Item("ED2") = vendor2EDSDSLookupDateTimePicker.Value
        End If

        'Add the row to the dataset
        ds.Tables("SDS").Rows.Add(dsNewRow)

        'Update the database using the data adapter
        da.Update(ds, "SDS")

        'MsgBox("Database is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database is now Closed")

    End Sub

    Private Sub addlog()

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
        Sql = "SELECT * FROM SDS_Log"

        'Send the search to the data adapter
        da = New OleDb.OleDbDataAdapter(Sql, con)

        'Tell the data adapter to fill the dataset
        da.Fill(ds, "SDSLog")

        'Add a new row to the dataset
        'Create a command builder 
        Dim cb As New OleDb.OleDbCommandBuilder(da)
        Dim dsNewRow As DataRow

        'Create new row in the dataset
        dsNewRow = ds.Tables("SDSLog").NewRow()

        'Add the new tag to the new datarow in the dataset
        dsNewRow.Item("Material") = materialSDSLookupComboBox.Text
        dsNewRow.Item("DateLogged") = Date.Today

        'Add the row to the dataset
        ds.Tables("SDSLog").Rows.Add(dsNewRow)

        'Update the database using the data adapter
        da.Update(ds, "SDSLog")

        'MsgBox("Database is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database is now Closed")
        Call clearForm()

    End Sub

    Private Sub materialSDSLookupComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles materialSDSLookupComboBox.SelectedIndexChanged
        If materialSDSLookupComboBox.Text <> "Enter Material" Then Call getMaterialInformation()
    End Sub

    Private Sub ExportSDSLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportSDSLogToolStripMenuItem.Click

        'Run Access Report to Export SDS Log data
        Dim objAccApp As New Microsoft.Office.Interop.Access.Application() 'Instantiate Access Application Object
        Dim strAccReport As String = "SDS_Log_Report" 'Get Selected ListBox Item
        objAccApp.OpenCurrentDatabase("F:\QA\SAFETY\MSDS\SDS Log.accdb", False, "") 'Open Database

        If Not objAccApp.Visible = False Then 'Do Not Show Access Window(s)
            objAccApp.Visible = False
        End If

        '---------------change default printer or printer select box----------------------

        objAccApp.Visible = True
        objAccApp.DoCmd.OpenReport(strAccReport, Microsoft.Office.Interop.Access.AcView.acViewPreview, Type.Missing, Type.Missing, AcWindowMode.acWindowNormal, Type.Missing) 'Open Selected Report
        'objAccApp.DoCmd.PrintOut(AcPrintRange.acPrintAll, Type.Missing, Type.Missing, AcPrintQuality.acHigh, Type.Missing, Type.Missing) 'Print Report

        'MsgBox("Click OK after Export", MsgBoxStyle.OkOnly, "Export SDS Log")

        'objAccApp.CloseCurrentDatabase() 'Close Database
        'objAccApp.Quit()
        objAccApp = Nothing 'Release Resources

        Call sendEmail()

        'Open Windows Explorer to SDS folder for adding attachment to Email
        Process.Start("F:\QA\SAFETY\MSDS\")

    End Sub

    Dim filePrefix, fileName As String

    Private Sub sendEmail()

        Dim OutlookMessage As Outlook.MailItem
        Dim AppOutlook As New Outlook.Application

        Dim objNS As Outlook._NameSpace = AppOutlook.Session
        Dim objFolder As Outlook.MAPIFolder
        objFolder = objNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)

        'Get recipients
        Dim recipients As String = "The NFL"

        'Create subject
        Dim subject As String = Nothing
        subject = "Latest SDSs"

        'Create HTML Body
        Dim MsgTxt As String = Nothing
        MsgTxt = MsgTxt &
        "<p>Hello Everyone,</p>" &
        "<p>Here are the highlights of the latest Safety Data Sheets (SDSs)." &
        "<br>Please contact me if you have any questions.</p>"

        'Include link to the HTML file

        'Closing
        MsgTxt = MsgTxt &
        "<p>Thank you,"

        'Signature location for sending via Document Control
        Dim sigstring As String = Nothing
        sigstring = Environ("appdata") & "\Microsoft\Signatures\Mario Valenzuela.htm"

        'Set variables for the signature
        Dim fso As Object = Nothing
        Dim ts As Object = Nothing
        Dim vsignature As Object = Nothing

        fso = CreateObject("Scripting.FileSystemObject")
        ts = fso.GetFile(sigstring).OpenAsTextStream(1, -2)
        vsignature = ts.readall

        Try
            OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            Dim Recipents As Outlook.Recipients = OutlookMessage.Recipients
            Recipents.Add(recipients)
            OutlookMessage.Subject = subject
            OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            OutlookMessage.HTMLBody = "<html><body>" & MsgTxt & vsignature & "</body></html>"
            'OutlookMessage.Save()
            OutlookMessage.Display()
            'OutlookMessage.Move(objFolder)
        Catch ex As Exception
            MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line    
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try
    End Sub

End Class
