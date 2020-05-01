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

    Private Function SourceFormatToLang(s As String, lang As String) As String
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
        StartSourceFormat()
    End Sub

    Private Sub cmbLangs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLangs.SelectedIndexChanged
        StartSourceFormat()
    End Sub

    Private Sub chkAutoDetectLang_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoDetectLang.CheckedChanged
        If chkAutoDetectLang.Checked Then StartSourceFormat()
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        My.Computer.Clipboard.SetText(txtOutput.Text)
    End Sub
    Private Sub StartSourceFormat()

        If chkAutoDetectLang.Checked Then
            cmbLangs.Text = GuessLanguage(txtInput.Text)
        End If
        txtOutput.Text = SourceFormatToLang(txtInput.Text, cmbLangs.Text)
    End Sub
    Private Function GuessLanguage(s As String) As String
        'points are accumulated for indicators of each language
        'this is a heuristic approach, not parsing so will not work for string literals or unusually formatted code

        Dim clikePts As Integer = 0
        Dim jsonPts As Integer = 0
        Dim pythonPts As Integer = 0
        Dim sqlPts As Integer = 0
        Dim vbnetPts As Integer = 0
        Const StrongIndicator As Integer = 10


        'c has many {} symbols
        clikePts += CountInstancesOfSubstrings(s, "{")
        clikePts += CountInstancesOfSubstrings(s, "}")

        'and ; on end of most lines
        clikePts += CountInstancesOfSubstrings(s, ";" & vbCrLf)

        '#define is common
        clikePts += CountInstancesOfSubstrings(s, "#define") * StrongIndicator

        Debug.Print("clikePts " & clikePts)

        'vbnet has End If, End Sub, End Function
        vbnetPts += CountInstancesOfSubstrings(s, "End If")
        vbnetPts += CountInstancesOfSubstrings(s, "End Sub")
        vbnetPts += CountInstancesOfSubstrings(s, "End Function")
        vbnetPts += CountInstancesOfSubstrings(s, " Then")

        Debug.Print("vbnetPts " & vbnetPts)

        'python uses def for functions
        pythonPts += CountInstancesOfSubstrings(s, "def ") * 10

        'has : on end of some lines
        pythonPts += CountInstancesOfSubstrings(s, ":" & vbCrLf)

        'python uses # for comments, but need to not count c's #define
        pythonPts += CountInstancesOfSubstrings(s, "#")
        pythonPts -= CountInstancesOfSubstrings(s, "#define")

        Debug.Print("pythonPts " & pythonPts)

        'json starts with {, ends with }
        If s.StartsWith("{") Then jsonPts += 5
        If s.EndsWith("}") Then jsonPts += 5

        'json lines tend to end with ,
        jsonPts += CountInstancesOfSubstrings(s, "," & vbCrLf)

        Debug.Print("jsonPts " & jsonPts)

        'sql contains keywords
        Dim su As String = s.ToUpper
        sqlPts += CountInstancesOfSubstrings(s, "SELECT ")
        sqlPts += CountInstancesOfSubstrings(s, " FROM ")
        sqlPts += CountInstancesOfSubstrings(s, " WHERE ")
        sqlPts += CountInstancesOfSubstrings(s, " JOIN ")
        sqlPts += CountInstancesOfSubstrings(s, "INSERT ")
        sqlPts += CountInstancesOfSubstrings(s, "UPDATE ")
        sqlPts += CountInstancesOfSubstrings(s, "DELETE ")
        sqlPts += CountInstancesOfSubstrings(s, " GROUP BY ")
        sqlPts += CountInstancesOfSubstrings(s, " HAVING ")
        sqlPts += CountInstancesOfSubstrings(s, " ORDER BY ")
        sqlPts += CountInstancesOfSubstrings(s, " COUNT ")
        sqlPts += CountInstancesOfSubstrings(s, " DISTINCT ")

        Debug.Print("sqlPts " & sqlPts)

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

    Public Function CountInstancesOfSubstrings(s As String, substr As String) As Integer
        Dim c As Integer = 0
        For i As Integer = 0 To s.Length - substr.Length
            If s.Substring(i, substr.Length) = substr Then
                c += 1
            End If
        Next i
        Return c
    End Function

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtInput.Text = String.Empty
    End Sub
End Class
