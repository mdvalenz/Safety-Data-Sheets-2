<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class vendorForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(vendorForm))
        Me.vendorNameTextBox = New System.Windows.Forms.TextBox()
        Me.saveSettingsButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'vendorNameTextBox
        '
        Me.vendorNameTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vendorNameTextBox.Location = New System.Drawing.Point(12, 12)
        Me.vendorNameTextBox.Name = "vendorNameTextBox"
        Me.vendorNameTextBox.Size = New System.Drawing.Size(420, 26)
        Me.vendorNameTextBox.TabIndex = 67
        '
        'saveSettingsButton
        '
        Me.saveSettingsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.saveSettingsButton.Location = New System.Drawing.Point(438, 10)
        Me.saveSettingsButton.Name = "saveSettingsButton"
        Me.saveSettingsButton.Size = New System.Drawing.Size(125, 30)
        Me.saveSettingsButton.TabIndex = 70
        Me.saveSettingsButton.TabStop = False
        Me.saveSettingsButton.Text = "Save and Exit"
        Me.saveSettingsButton.UseVisualStyleBackColor = True
        '
        'vendorForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(574, 53)
        Me.Controls.Add(Me.saveSettingsButton)
        Me.Controls.Add(Me.vendorNameTextBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "vendorForm"
        Me.Text = "Add Vendor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents vendorNameTextBox As TextBox
    Friend WithEvents saveSettingsButton As Button
End Class
