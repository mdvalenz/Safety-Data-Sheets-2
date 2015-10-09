Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Access

Public Class mainForm

    'Dim materialID As Integer = Nothing
    Dim materialName As String = Nothing
    Dim materialID As String = Nothing
    Dim newMaterial As Boolean = Nothing
    Dim newVendor1 As Boolean = Nothing
    Dim newVendor2 As Boolean = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Created by Mario Valenzuela
        Call loadForm()
    End Sub

    Private Sub loadForm()
        'TODO: This line of code loads data into the 'SDS_LogDataSet1.SDS_Log' table. You can move, or remove it, as needed.
        Me.SDS_LogTableAdapter.Fill(Me.SDS_LogDataSet1.SDS_Log)
        'TODO: This line of code loads data into the 'SDS_LogDataSet.SDS_List' table. You can move, or remove it, as needed.
        Me.SDS_ListTableAdapter.Fill(Me.SDS_LogDataSet.SDS_List)
        'TODO: This line of code loads data into the 'SDS_LogDataSet.VendorList' table. You can move, or remove it, as needed.
        Me.VendorListTableAdapter.Fill(Me.SDS_LogDataSet.VendorList)
        'TODO: This line of code loads data into the 'SDS_LogDataSet2.VendorList1' table. You can move, or remove it, as needed.
        Me.VendorList1TableAdapter.Fill(Me.SDS_LogDataSet2.VendorList1)

        'Fill all datagridviews and dropdown lists
        Call loadDefaultData()

        My.Settings.materialName = ""

        materialSDSLookupComboBox.Text = "Enter material"
        SDSLogComboBox.Text = "Enter material"
        vendor1ComboBox.Text = "Enter vendor"
        vendor2ComboBox.Text = "Enter vendor"
        materialSDSLookupComboBox.Focus()

    End Sub

    Private Sub loadDefaultData()

        'Fill DataGridViews with the SDS List and SDS Log
        Call loadSDSLog()

        'Set dates
        vendor1EDSDSLookupDateTimePicker.Value = Date.Today
        vendor2EDSDSLookupDateTimePicker.Value = Date.Today

        'Fill the Material List dropdown from the database
        'Call loadMaterials()

    End Sub

    Private Sub SettingsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem1.Click
        settingsForm.Show()
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
        My.Settings.materialName = SDSLogComboBox.Text
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

                vendor1ComboBox.Text = dr.Item(4).ToString
                If vendor1ComboBox.Text <> "" Then vendor1EDSDSLookupDateTimePicker.Value = dr.Item(5)
                vendor2ComboBox.Text = dr.Item(6).ToString
                If vendor2ComboBox.Text <> "" Then vendor2EDSDSLookupDateTimePicker.Value = dr.Item(7)

                'End While

                'MsgBox("Database is now Closed")
                newMaterial = False
            Else
                '    Dim inputResult As Integer = MsgBox("Material Not Found. Would you like to add a new material?", MsgBoxStyle.YesNo, "Add New Material?")

                '    If inputResult = 6 Then

                'Clear the rest of the form
                CASNumberSDSLookupTextBox.Text = ""
                hazardsSDSLookupTextBox.Text = ""
                vendor1ComboBox.Text = ""
                vendor1EDSDSLookupDateTimePicker.Value = Date.Today
                vendor2ComboBox.Text = ""
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

    Private Sub checkMaterial()

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
        Sql = "SELECT ID, Material FROM SDS_List WHERE Material=@materialName"

        Dim cm As New OleDb.OleDbCommand(Sql, con)
        cm.Parameters.AddWithValue("@materialName", materialName)

        'Send the search to the data reader
        Dim dr As OleDb.OleDbDataReader = cm.ExecuteReader

        Try
            If dr.Read() Then
                newMaterial = False
                My.Settings.materialID = dr.Item(0).ToString
            Else
                newMaterial = True
            End If

        Finally

            'Closing the database connection
            dr.Close()
            con.Close()

        End Try

    End Sub

    Private Sub checkVendors()

        'Get Material Name
        Dim vendor1Name As String = vendor1ComboBox.Text
        Dim vendor2Name As String = vendor2ComboBox.Text

        'Search for vendor in Vendor list
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

        If vendor1Name <> "" Then

            'Search string to select all tags from the tagTable
            Sql = "SELECT * FROM VendorList WHERE VendorName=@vendorName"

            Dim cm As New OleDb.OleDbCommand(Sql, con)
            cm.Parameters.AddWithValue("@vendorName", vendor1Name)

            'Send the search to the data reader
            Dim dr As OleDb.OleDbDataReader = cm.ExecuteReader

            Try
                If dr.Read() Then
                    newVendor1 = False
                Else
                    newVendor1 = True
                End If

            Finally

            End Try
            dr.Close()

        End If

        If vendor2Name <> "" Then

            'Search string to select all tags from the tagTable
            Sql = "SELECT * FROM VendorList WHERE VendorName=@vendorName"

            Dim cm As New OleDb.OleDbCommand(Sql, con)
            cm.Parameters.AddWithValue("@vendorName", vendor2Name)

            'Send the search to the data reader
            Dim dr As OleDb.OleDbDataReader = cm.ExecuteReader

            Try
                If dr.Read() Then
                    newVendor2 = False
                Else
                    newVendor2 = True
                End If

            Finally

            End Try
            dr.Close()

        End If

        'Closing the database connection
        con.Close()

    End Sub

    Private Sub clearForm()

        CASNumberSDSLookupTextBox.Text = ""
        hazardsSDSLookupTextBox.Text = ""
        vendor1ComboBox.Text = ""
        vendor1EDSDSLookupDateTimePicker.Value = Date.Today
        vendor2ComboBox.Text = ""
        vendor2EDSDSLookupDateTimePicker.Value = Date.Today

        materialSDSLookupComboBox.Text = "Enter material"
        SDSLogComboBox.Text = "Enter material"
        vendor1ComboBox.Text = "Enter vendor"
        vendor2ComboBox.Text = "Enter vendor"

        materialSDSLookupComboBox.Focus()

    End Sub

    Private Sub clearSDSLookupButton_Click(sender As Object, e As EventArgs) Handles clearSDSLookupButton.Click
        Call clearForm()
        Call hazardsForm.clearAllChecks()
    End Sub

    Private Sub resetSDSLookupButton_Click(sender As Object, e As EventArgs) Handles resetSDSLookupButton.Click
        Call getMaterialInformation()
        Call hazardsForm.clearAllChecks()
    End Sub

    Private Sub saveSDSLookupButton_Click(sender As Object, e As EventArgs) Handles saveSDSLookupButton.Click

        'Check if required information is blank
        If materialSDSLookupComboBox.Text = "" Or CASNumberSDSLookupTextBox.Text = "" Or hazardsSDSLookupTextBox.Text = "" Or vendor1ComboBox.Text = "" Then
            MsgBox("Please enter all required information.")
            Exit Sub
        End If

        If vendor1ComboBox.Text = "Enter vendor" Then
            MsgBox("Please enter a vendor", MsgBoxStyle.Exclamation, "Enter Vendor")
            Exit Sub
        End If

        Call checkMaterial()
        Call checkVendors()

        If newVendor1 = True Then
            Call saveVendor("Vendor1")
        End If

        If newVendor2 = True Then
            Call saveVendor("Vendor2")
        End If

        If newMaterial = False Then
            My.Settings.materialName = materialSDSLookupComboBox.Text
            Call editMaterial()
        Else
            Call saveMaterial()
        End If

        Call addlog()
        'Call loadMaterials()
        Call loadSDSLog()
        Call hazardsForm.clearAllChecks()
        Call refreshMaterials()

    End Sub

    Private Sub refreshMaterials()
        Try
            Me.Controls.Clear() 'removes all the controls on the form
        Catch
        End Try

        InitializeComponent() 'load all the controls again
        vendor2ComboBox.DataSource = VendorListBindingSource2

        loadForm() 'loads from load items

    End Sub

    Private Sub editMaterial()

        'Get Variables
        materialName = My.Settings.materialName
        materialID = My.Settings.materialID

        Dim hazards As String = hazardsSDSLookupTextBox.Text
        Dim CASNumber As String = CASNumberSDSLookupTextBox.Text
        Dim vendor1 As String = vendor1ComboBox.Text
        Dim vendor1Date As Date = Format(vendor1EDSDSLookupDateTimePicker.Value, "M/d/yyyy")
        Dim vendor2 As String = vendor2ComboBox.Text
        Dim vendor2Date As Date = Format(vendor2EDSDSLookupDateTimePicker.Value, "M/d/yyyy")

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

        Try
            'Search string to select all columns from the table
            Sql = "SELECT * FROM SDS_List WHERE ID=" & materialID

            'Send the search to the data adapter
            da = New OleDb.OleDbDataAdapter(Sql, con)

            'Tell the data adapter to fill the dataset
            da.Fill(ds, "SDS")

            'Edit row in the dataset
            'Create a command builder 
            Dim cb As New OleDb.OleDbCommandBuilder(da)

            'MsgBox(ds.Tables(0).Rows(0).Item(0) & ", " & ds.Tables(0).Rows(0).Item(1))
            'MsgBox(ds.Tables(0).Rows(0).Item(2) & ", " & ds.Tables(0).Rows(0).Item(3))

            'Add the new tag to the new datarow in the dataset
            ds.Tables(0).Rows(0).Item(2) = hazardsSDSLookupTextBox.Text
            ds.Tables(0).Rows(0).Item(3) = CASNumberSDSLookupTextBox.Text

            If vendor1ComboBox.Text <> "" Then

                'MsgBox(ds.Tables(0).Rows(0).Item(4) & ", " & ds.Tables(0).Rows(0).Item(5))

                ds.Tables(0).Rows(0).Item(4) = vendor1ComboBox.Text
                ds.Tables(0).Rows(0).Item(5) = vendor1EDSDSLookupDateTimePicker.Value

                'MsgBox(ds.Tables(0).Rows(0).Item(4) & ", " & ds.Tables(0).Rows(0).Item(5))

            End If

            If vendor2ComboBox.Text <> "" Then

                'MsgBox(ds.Tables(0).Rows(0).Item(6) & ", " & ds.Tables(0).Rows(0).Item(7))

                ds.Tables(0).Rows(0).Item(6) = vendor2ComboBox.Text
                ds.Tables(0).Rows(0).Item(7) = vendor2EDSDSLookupDateTimePicker.Value

                'MsgBox(ds.Tables(0).Rows(0).Item(6) & ", " & ds.Tables(0).Rows(0).Item(7))

            End If

            'Update the database using the data adapter
            da.Update(ds, "SDS")

        Catch ex As Exception
            MsgBox("An exception occurred:" & vbCrLf & ex.Message)
        End Try

        'MsgBox("Database is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database is now Closed")

    End Sub

    Private Sub saveMaterial()

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

        'Search string to select all columns from the table
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

        If vendor1ComboBox.Text <> "" Then
            dsNewRow.Item("Vendor1") = vendor1ComboBox.Text
            dsNewRow.Item("ED1") = vendor1EDSDSLookupDateTimePicker.Value
        End If

        If vendor2ComboBox.Text <> "" Then
            dsNewRow.Item("Vendor2") = vendor2ComboBox.Text
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

    Private Sub saveVendor(vendorNumber As String)

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

        'Search string to select all columns from the table
        Sql = "SELECT * FROM VendorList"

        'Send the search to the data adapter
        da = New OleDb.OleDbDataAdapter(Sql, con)

        'Tell the data adapter to fill the dataset
        da.Fill(ds, "Vendor")

        'Add a new row to the dataset
        'Create a command builder 
        Dim cb As New OleDb.OleDbCommandBuilder(da)
        Dim dsNewRow As DataRow

        'Create new row in the dataset
        dsNewRow = ds.Tables("Vendor").NewRow()

        Select Case vendorNumber
            Case "Vendor1"
                'Add the new tag to the new datarow in the dataset
                dsNewRow.Item("VendorName") = vendor1ComboBox.Text
            Case "Vendor2"
                'Add the new tag to the new datarow in the dataset
                dsNewRow.Item("VendorName") = vendor2ComboBox.Text
        End Select

        'Add the row to the dataset
        ds.Tables("Vendor").Rows.Add(dsNewRow)

        'Update the database using the data adapter
        da.Update(ds, "Vendor")

        'MsgBox("Database Is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database Is now Closed")

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
        Sql = "Select * FROM SDS_Log"

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

        'MsgBox("Database Is now open")

        'Closing the database connection
        con.Close()

        'MsgBox("Database Is now Closed")
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

        Try
            objAccApp.DoCmd.OpenReport(strAccReport, Microsoft.Office.Interop.Access.AcView.acViewPreview, Type.Missing, Type.Missing, AcWindowMode.acWindowNormal, Type.Missing) 'Open Selected Report
            'objAccApp.DoCmd.PrintOut(AcPrintRange.acPrintAll, Type.Missing, Type.Missing, AcPrintQuality.acHigh, Type.Missing, Type.Missing) 'Print Report

            'MsgBox("Click OK after Export", MsgBoxStyle.OkOnly, "Export SDS Log")

            'objAccApp.CloseCurrentDatabase() 'Close Database
            'objAccApp.Quit()
            objAccApp = Nothing 'Release Resources

            Call sendEmail()

            'Open Windows Explorer to SDS folder for adding attachment to Email
            Process.Start("F:\QA\SAFETY\MSDS\")
        Catch
        End Try

    End Sub

    Dim filePrefix, fileName As String

    Private Sub selectHazardsButton_Click(sender As Object, e As EventArgs) Handles selectHazardsButton.Click
        hazardsForm.ShowDialog()
    End Sub

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
