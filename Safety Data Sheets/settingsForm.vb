Public Class settingsForm

    Private Sub browseLocationSettingsButton_Click(sender As Object, e As EventArgs) Handles browseLocationSettingsButton.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "F:\QA\SAFETY\SDS\"
        fd.Filter = "Microsoft Access Database (*.accdb)|*.accdb"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            My.Settings.SDSDBLocation = fd.FileName
            DBLocationSettingsTextBox.Text = My.Settings.SDSDBLocation
        End If
    End Sub

    Private Sub saveSettingsButton_Click(sender As Object, e As EventArgs) Handles saveSettingsButton.Click
        Close()
    End Sub

End Class