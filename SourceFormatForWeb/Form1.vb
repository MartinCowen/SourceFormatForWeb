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

        'add the code start tag
        r = r.Insert(0, "<code class=""language-" & lang & """>")

        'want to put the end tag on the same line to avoid extra line breaks
        r = r.TrimEnd(vbCrLf)

        'close the tag
        r = r & "</code>"
        Return r
    End Function

    Private Sub txtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged
        txtOutput.Text = Convert(txtInput.Text, cmbLangs.Text)
    End Sub

    Private Sub cmbLangs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLangs.SelectedIndexChanged
        txtOutput.Text = Convert(txtInput.Text, cmbLangs.Text)
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        My.Computer.Clipboard.SetText(txtOutput.Text)
    End Sub

    Private Function GuessLanguage(s As String) As String
        'points are accumulated for indicators of each language
        'this is a heuristic approach, not parsing so will not work for string literals or unusually formatted code

        Dim clikePts As Integer
        Dim jsonPts As Integer
        Dim pythonPts As Integer
        Dim sqlPts As Integer
        Dim vbnetPts As Integer

        'c has many {} symbols
        clikePts += Len(s) - Len(s.Replace("{", ""))
        clikePts += Len(s) - Len(s.Replace("}", ""))

        'and ; on end of most lines
        clikePts += s.Split(";" & vbCrLf).Length - 1

        '#define is common
        clikePts += s.Split("#define").Length - 1 + 10

        'vbnet has End If, End Sub, End Function
        vbnetPts += s.Split("End If").Length - 1
        vbnetPts += s.Split("End Sub").Length - 1
        vbnetPts += s.Split("End Function").Length - 1

        'python uses def for functions
        pythonPts += s.Split("def ").Length - 1 + 10

        'has : on end of some lines
        pythonPts += s.Split(":" & vbCrLf).Length - 1

        'python uses # for comments, but need to not count c's #define
        pythonPts += s.Split("#").Length - 1
        pythonPts -= s.Split("#define").Length - 1

        'json starts with {, ends with }
        If s.StartsWith("{") Then jsonPts += 5
        If s.EndsWith("}") Then jsonPts += 5

        'json lines tend to end with ,
        clikePts += s.Split("," & vbCrLf).Length - 1

        'sql contains keywords
        Dim su As String = s.ToUpper
        sqlPts += su.Split("SELECT").Length - 1
        sqlPts += su.Split("FROM").Length - 1
        sqlPts += su.Split("WHERE").Length - 1
        sqlPts += su.Split("JOIN").Length - 1
        sqlPts += su.Split("INSERT").Length - 1
        sqlPts += su.Split("UPDATE").Length - 1
        sqlPts += su.Split("DELETE").Length - 1
        sqlPts += su.Split("GROUP BY").Length - 1
        sqlPts += su.Split("HAVING").Length - 1
        sqlPts += su.Split("ORDER BY").Length - 1
        sqlPts += su.Split("COUNT").Length - 1
        sqlPts += su.Split("DISTINCT").Length - 1

        'decide which lang is most likely
        'using max score

        If clikePts > jsonPts AndAlso clikePts > pythonPts AndAlso clikePts > sqlPts AndAlso clikePts > vbnetPts Then
            Return "clike"
        End If

        If jsonPts > clikePts AndAlso jsonPts > pythonPts AndAlso jsonPts > sqlPts AndAlso jsonPts > vbnetPts Then
            Return "json"
        End If

        If pythonPts > clikePts AndAlso pythonPts > jsonPts AndAlso pythonPts > sqlPts AndAlso pythonPts > vbnetPts Then
            Return "python"
        End If

        If sqlPts > clikePts AndAlso sqlPts > jsonPts AndAlso sqlPts > pythonPts AndAlso sqlPts > vbnetPts Then
            Return "sql"
        End If

        If vbnetPts > clikePts AndAlso vbnetPts > jsonPts AndAlso vbnetPts > pythonPts AndAlso vbnetPts > sqlPts Then
            Return "vbnet"
        End If

        Return ""
    End Function
End Class
