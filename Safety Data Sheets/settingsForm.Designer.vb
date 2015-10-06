<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class settingsForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(settingsForm))
        Me.saveSettingsButton = New System.Windows.Forms.Button()
        Me.browseLocationSettingsButton = New System.Windows.Forms.Button()
        Me.DBLocationSettingsTextBox = New System.Windows.Forms.TextBox()
        Me.DBLocationSettingsLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'saveSettingsButton
        '
        Me.saveSettingsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.saveSettingsButton.Location = New System.Drawing.Point(397, 86)
        Me.saveSettingsButton.Name = "saveSettingsButton"
        Me.saveSettingsButton.Size = New System.Drawing.Size(125, 30)
        Me.saveSettingsButton.TabIndex = 69
        Me.saveSettingsButton.TabStop = False
        Me.saveSettingsButton.Text = "Save and Exit"
        Me.saveSettingsButton.UseVisualStyleBackColor = True
        '
        'browseLocationSettingsButton
        '
        Me.browseLocationSettingsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.browseLocationSettingsButton.Location = New System.Drawing.Point(442, 30)
        Me.browseLocationSettingsButton.Name = "browseLocationSettingsButton"
        Me.browseLocationSettingsButton.Size = New System.Drawing.Size(80, 30)
        Me.browseLocationSettingsButton.TabIndex = 68
        Me.browseLocationSettingsButton.Text = "Browse"
        Me.browseLocationSettingsButton.UseVisualStyleBackColor = True
        '
        'DBLocationSettingsTextBox
        '
        Me.DBLocationSettingsTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.Safety_Data_Sheets.My.MySettings.Default, "SDSDBLocation", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.DBLocationSettingsTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DBLocationSettingsTextBox.Location = New System.Drawing.Point(16, 32)
        Me.DBLocationSettingsTextBox.Name = "DBLocationSettingsTextBox"
        Me.DBLocationSettingsTextBox.Size = New System.Drawing.Size(420, 26)
        Me.DBLocationSettingsTextBox.TabIndex = 66
        Me.DBLocationSettingsTextBox.Text = Global.Safety_Data_Sheets.My.MySettings.Default.SDSDBLocation
        '
        'DBLocationSettingsLabel
        '
        Me.DBLocationSettingsLabel.AutoSize = True
        Me.DBLocationSettingsLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DBLocationSettingsLabel.Location = New System.Drawing.Point(12, 9)
        Me.DBLocationSettingsLabel.Name = "DBLocationSettingsLabel"
        Me.DBLocationSettingsLabel.Size = New System.Drawing.Size(144, 20)
        Me.DBLocationSettingsLabel.TabIndex = 67
        Me.DBLocationSettingsLabel.Text = "Database Location"
        '
        'settingsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(546, 135)
        Me.Controls.Add(Me.saveSettingsButton)
        Me.Controls.Add(Me.browseLocationSettingsButton)
        Me.Controls.Add(Me.DBLocationSettingsTextBox)
        Me.Controls.Add(Me.DBLocationSettingsLabel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "settingsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents saveSettingsButton As Button
    Friend WithEvents browseLocationSettingsButton As Button
    Friend WithEvents DBLocationSettingsTextBox As TextBox
    Friend WithEvents DBLocationSettingsLabel As Label
End Class
