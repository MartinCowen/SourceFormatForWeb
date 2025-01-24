Imports System.Net

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

    ''' <summary>
    ''' Formats source code string s by replacing tabs with 4 spaces, inserting pre and code tags and optionally HTML encoding
    ''' </summary>
    ''' <param name="s">input string to be formatted</param>
    ''' <param name="lang">name of language</param>
    ''' <param name="htmlencode">whether to apply HTML Encoding to escape angle brackets</param>
    ''' <returns>formatted string</returns>
    Private Function SourceFormatToLang(s As String, lang As String, htmlencode As Boolean) As String
        Dim r As String

        If htmlencode Then r = WebUtility.HtmlEncode(s) Else r = s

        'replace tabs with 4 spaces
        r = r.Replace(vbTab, "    ")

        'add the pre and code start tag
        r = r.Insert(0, "<pre><code class=""language-" & lang & """>")

        'want to put the end tag on the same line to avoid extra line breaks
        r = r.TrimEnd(vbCrLf)

        'close the tag
        r = r & "</code></pre>"
        Return r
    End Function

    Private Sub txtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged
        StartSourceFormat()
    End Sub

    Private Sub cmbLangs_TextChanged(sender As Object, e As EventArgs) Handles cmbLangs.TextChanged
        StartSourceFormat()
    End Sub

    Private Sub chkAutoDetectLang_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoDetectLang.CheckedChanged
        If chkAutoDetectLang.Checked Then StartSourceFormat()
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        Try
            My.Computer.Clipboard.SetText(txtOutput.Text)
        Catch ex As Exception
            MessageBox.Show("Could not copy to clipboard!")
        End Try
    End Sub
    Private Sub StartSourceFormat()

        If chkAutoDetectLang.Checked Then
            cmbLangs.Text = GuessLanguage(txtInput.Text)
        End If
        txtOutput.Text = SourceFormatToLang(txtInput.Text, cmbLangs.Text, chkHTMLEncode.Checked)
    End Sub
    ''' <summary>
    ''' Guesses programming language from a small predefined set using simple rules. Designed for snippets.
    ''' </summary>
    ''' <param name="s">input string</param>
    ''' <returns>name of language</returns>
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

        'and ; on end of most lines - an inline comment could mess this one up
        clikePts += CountInstancesOfSubstrings(s, ";" & vbCrLf)

        'often uses c++ style comments //
        clikePts += CountInstancesOfSubstrings(s, "//")

        'or the c style /* and */ pairs
        clikePts += CountInstancesOfSubstrings(s, "/*")
        clikePts += CountInstancesOfSubstrings(s, "*/")


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

    ''' <summary>
    ''' Counts instances of substrings within strings
    ''' </summary>
    ''' <param name="s">string to be searched</param>
    ''' <param name="substr">substring</param>
    ''' <returns>Number of instances</returns>
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
        txtInput.Focus()
    End Sub

    Private Sub chkHTMLEncode_CheckedChanged(sender As Object, e As EventArgs) Handles chkHTMLEncode.CheckedChanged
        StartSourceFormat()
    End Sub



End Class
