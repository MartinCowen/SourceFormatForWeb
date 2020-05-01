Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateLangs()
    End Sub
    Private Sub PopulateLangs()
        With cmbLangs
            .Enabled = False
            For Each s As String In My.Settings.Language
                cmbLangs.Items.Add(s)
            Next
            .Enabled = True

        End With
    End Sub

    Private Function Convert(s As String, lang As String) As String
        Dim r As String
        'replace tabs with 4 spaces
        r = s.Replace(vbTab, "    ")

        r = r.Insert(0, "<code class=""language-" & lang & """>")

        r = r.TrimEnd(vbCrLf)
        r = r & "</code>"
        Return r
    End Function

    Private Sub txtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged

        txtOutput.Text = Convert(txtInput.Text, cmbLangs.Text)
    End Sub

    Private Sub cmbLangs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLangs.SelectedIndexChanged
        txtOutput.Text = Convert(txtInput.Text, cmbLangs.Text)
    End Sub
End Class
