Imports System.Text.RegularExpressions

Public Class Form1

    Dim txt As String
    Dim enabledEvents As Boolean

    'Import from File
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = "Open A File..."
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "Text Files|*.txt"
        If OpenFileDialog1.ShowDialog Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox1.Text = OpenFileDialog1.FileName
            TextBox2.Text = My.Computer.FileSystem.ReadAllText(TextBox1.Text)
            txt = TextBox2.Text
        End If
        enabledEvents = True
    End Sub

    'RefreshFilters
    Public Sub Refreshfilters()
        Call filterF1()
        Call filterF2()
        Call filterF3()
        Call filterF6()
        Call filterF7()
        Call filterF8()
        Call filterF5()
        Call filterF4()
    End Sub

    '1 None Tab Comma Other
    Public Sub filterF1()
        txt = TextBox2.Text
        If (RadioButton11.Checked = True) Then
            TextBox3.Text = txt
        Else
            If (RadioButton5.Checked = True) Then
                TextBox3.Text = Join(Split(txt, vbTab), vbNewLine)
            Else
                If (RadioButton3.Checked = True) Then
                    TextBox3.Text = Join(Split(txt, ","), vbNewLine)
                Else
                    TextBox3.Text = Join(Split(txt, TextBox17.Text), vbNewLine)
                End If
            End If
        End If
    End Sub

    'None
    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton11.CheckedChanged
        If enabledEvents Then
            Call Refreshfilters()
        End If
    End Sub

    'Tab
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        Call Refreshfilters()
    End Sub

    'Comma
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Call Refreshfilters()
    End Sub

    'Text
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Call Refreshfilters()
    End Sub


    Private Sub TextBox17_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        If (Len(TextBox17.Text) > 0) Then
            RadioButton4.Enabled = True
            RadioButton4.Checked = True
            RadioButton4.Select()
        Else
            RadioButton4.Enabled = False
            RadioButton4.Checked = False
            RadioButton11.Select()
        End If
        TextBox17.Select()
    End Sub


    '2 Contain1 if not blank
    Public Sub filterF2()
        If (Len(TextBox4.Text) = 0) Then
            TextBox6.Enabled = False
            TextBox8.Enabled = False
            RadioButton10.Enabled = False
            RadioButton9.Enabled = False
            RadioButton7.Enabled = False
            RadioButton8.Enabled = False
        Else
            txt = TextBox3.Text
            TextBox3.Clear()
            If (Len(TextBox6.Text) = 0 Or TextBox6.Enabled = False) Then
                TextBox8.Enabled = False
                'TextBox3.Text = txt
                RadioButton10.Enabled = False
                RadioButton9.Enabled = False

                'Only4
                If enabledEvents Then
                    Dim TextLines() As String = txt.Split(vbNewLine)
                    For Each lineTxt In TextLines
                        If (InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0) Then
                            If Len(TextBox3.Text) = 0 Then
                                TextBox3.Text = lineTxt
                            Else
                                TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                            End If
                        End If
                    Next
                    enabledEvents = False
                    RadioButton7.Enabled = True
                    RadioButton8.Enabled = True
                    RadioButton8.Select()
                    TextBox6.Enabled = True
                    enabledEvents = True
                End If

            Else


                If (Len(TextBox8.Text) = 0 Or TextBox8.Enabled = False) Then

                    If (RadioButton7.Checked = True) Then
                        '4AND6
                        If enabledEvents Then
                            Dim TextLines() As String = txt.Split(vbNewLine)
                            For Each lineTxt In TextLines
                                If (InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 And InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) Then
                                    If Len(TextBox3.Text) = 0 Then
                                        TextBox3.Text = lineTxt
                                    Else
                                        TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                                    End If
                                End If
                            Next
                            enabledEvents = False
                            RadioButton9.Enabled = True
                            RadioButton10.Enabled = True
                            TextBox8.Enabled = True
                            enabledEvents = True
                            If (RadioButton9.Checked = False And RadioButton10.Checked = False) Then
                                enabledEvents = False
                                RadioButton9.Select()
                                enabledEvents = True
                            End If
                        End If

                    Else
                        '4OR6
                        If enabledEvents Then
                            If (RadioButton8.Checked = True) Then
                                Dim TextLines() As String = txt.Split(vbNewLine)
                                For Each lineTxt In TextLines
                                    If (InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 Or InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) Then
                                        If Len(TextBox3.Text) = 0 Then
                                            TextBox3.Text = lineTxt
                                        Else
                                            TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                                        End If
                                    End If
                                Next
                                enabledEvents = False
                                RadioButton9.Enabled = True
                                RadioButton10.Enabled = True
                                TextBox8.Enabled = True
                                enabledEvents = True
                            End If
                            If (RadioButton9.Checked = False And RadioButton10.Checked = False) Then
                                enabledEvents = False
                                RadioButton9.Select()
                                enabledEvents = True
                            End If
                        End If
                    End If


                Else

                    If (RadioButton10.Checked = True) Then

                        If enabledEvents Then
                            Dim TextLines() As String = txt.Split(vbNewLine)

                            If (RadioButton7.Checked = True) Then
                                '4AND6AND8
                                For Each lineTxt In TextLines
                                    If ((InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 And InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) And InStr(lineTxt, TextBox8.Text, vbTextCompare) > 0) Then
                                        If Len(TextBox3.Text) = 0 Then
                                            TextBox3.Text = lineTxt
                                        Else
                                            TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                                        End If
                                    End If
                                Next
                            Else
                                '4OR6AND8
                                For Each lineTxt In TextLines
                                    If ((InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 Or InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) And InStr(lineTxt, TextBox8.Text, vbTextCompare) > 0) Then
                                        If Len(TextBox3.Text) = 0 Then
                                            TextBox3.Text = lineTxt
                                        Else
                                            TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                                        End If
                                    End If
                                Next
                            End If
                        End If

                    Else

                        If enabledEvents Then
                            Dim TextLines() As String = txt.Split(vbNewLine)

                            If (RadioButton7.Checked = True) Then
                                '4AND6OR8
                                For Each lineTxt In TextLines
                                    If ((InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 And InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) Or InStr(lineTxt, TextBox8.Text, vbTextCompare) > 0) Then
                                        If Len(TextBox3.Text) = 0 Then
                                            TextBox3.Text = lineTxt
                                        Else
                                            TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                                        End If
                                    End If
                                Next
                            Else
                                '4OR6OR8
                                For Each lineTxt In TextLines
                                    If ((InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 Or InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) Or InStr(lineTxt, TextBox8.Text, vbTextCompare) > 0) Then
                                        If Len(TextBox3.Text) = 0 Then
                                            TextBox3.Text = lineTxt
                                        Else
                                            TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                                        End If
                                    End If
                                Next
                            End If

                        End If
                    End If

                End If


            End If

        End If

    End Sub



    Public Sub filterF3()
        'Start With
        If (Len(TextBox10.Text) <> 0 Or Len(TextBox11.Text) <> 0) Then
            If (Len(TextBox11.Text) <> 0) Then
                txt = TextBox3.Text
                TextBox3.Clear()
                Dim TextLines() As String = txt.Split(vbNewLine)
                For Each lineTxt In TextLines
                    If (Microsoft.VisualBasic.Left(Trim(Replace(lineTxt, Chr(10), "")), Len(TextBox11.Text)) = TextBox11.Text) Then
                        If Len(TextBox3.Text) = 0 Then
                            TextBox3.Text = Trim(lineTxt)
                        Else
                            TextBox3.Text = TextBox3.Text & vbNewLine & Trim(lineTxt)
                        End If
                    End If
                Next
            End If

            'End With
            If (Len(TextBox10.Text) <> 0) Then
                txt = TextBox3.Text
                TextBox3.Clear()
                Dim TextLines() As String = txt.Split(vbNewLine)
                For Each lineTxt In TextLines
                    If (Microsoft.VisualBasic.Right(Trim(Replace(lineTxt, Chr(10), "")), Len(TextBox10.Text)) = TextBox10.Text) Then
                        If Len(TextBox3.Text) = 0 Then
                            TextBox3.Text = Trim(lineTxt)
                        Else
                            TextBox3.Text = TextBox3.Text & vbNewLine & Trim(lineTxt)
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    'SpecificOccurance
    Public Sub filterF4()
        If (Len(TextBox15.Text) > 0) Then
            If (TextBox15.Text > 0) Then
                txt = TextBox3.Text
                TextBox3.Clear()
                Dim TextLines() As String = txt.Split(vbNewLine)
                TextBox3.Text = TextLines(TextBox15.Text - 1)
            End If
        End If
    End Sub

    'Word at position
    Public Sub filterF5()
        If (Len(TextBox5.Text) > 0) Then
            If (TextBox5.Text > 0) Then
                txt = TextBox3.Text
                TextBox3.Clear()
                Dim TextLines() As String = txt.Split(vbNewLine)
                Dim linewords() As String
                Dim wordBrkChar As String

                If (Len(TextBox9.Text) > 0) Then
                    wordBrkChar = TextBox9.Text
                Else
                    wordBrkChar = " "
                End If

                For Each lineTxt In TextLines
                    linewords = lineTxt.Split(wordBrkChar)
                    On Error Resume Next
                    If Len(TextBox3.Text) = 0 Then
                        TextBox3.Text = Trim(linewords(TextBox5.Text - 1))
                    Else
                        TextBox3.Text = TextBox3.Text & vbNewLine & Trim(linewords(TextBox5.Text - 1))
                    End If
                    If Err.Number <> 0 Then
                        Err.Clear()
                    End If
                Next
            End If
        End If
    End Sub


    'EmailFilter
    Public Sub filterF6()
        If RadioButton6.Checked = True Then
            Dim mc As MatchCollection
            Dim i As Integer
            txt = TextBox3.Text
            TextBox3.Clear()
            mc = Regex.Matches(txt, "([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})")
            For i = 0 To (mc.Count - 1)
                If Len(TextBox3.Text) = 0 Then
                    TextBox3.Text = mc(i).Value
                Else
                    TextBox3.Text = TextBox3.Text & vbNewLine & mc(i).Value
                End If
            Next

            txt = TextBox3.Text
            TextBox3.Clear()
            Dim TextLines() As String = txt.Split(vbNewLine).Distinct().ToArray()

            For Each lineTxt In TextLines

                If Len(TextBox3.Text) = 0 Then
                    TextBox3.Text = lineTxt
                Else
                    TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                End If

            Next
        End If

    End Sub

    'Mobile Numbers
    Public Sub filterF7()
        If RadioButton13.Checked = True Then
            Dim mc As MatchCollection
            Dim i As Integer

            txt = TextBox3.Text
            TextBox3.Clear()
            mc = Regex.Matches(txt, "([+]{0,1}[1-9]{1}[0-9]{1}|0|)[-]{0,1}?[1-9]{1}[0-9]{9}")
            For i = 0 To (mc.Count - 1)
                If Len(TextBox3.Text) = 0 Then
                    TextBox3.Text = mc(i).Value
                Else
                    TextBox3.Text = TextBox3.Text & vbNewLine & mc(i).Value
                End If
            Next


            txt = TextBox3.Text
            TextBox3.Clear()
            Dim TextLines() As String = txt.Split(vbNewLine).Distinct().ToArray()

            For Each lineTxt In TextLines

                If Len(TextBox3.Text) = 0 Then
                    TextBox3.Text = lineTxt
                Else
                    TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                End If
            Next
        End If
    End Sub


    'All Phone Numbers
    Public Sub filterF8()
        If RadioButton14.Checked = True Then
            Dim mc As MatchCollection
            Dim mc1 As MatchCollection
            Dim i As Integer
            txt = TextBox3.Text
            TextBox3.Clear()
            mc = Regex.Matches(txt, "[0-9]{2}[-]{0,1}[0-9]{8}|[0-9]{3,4}[-]{0,1}[0-9]{6,7}|[0-9()\-.]{8,}")
            mc1 = Regex.Matches(txt, "([+]{0,1}[1-9]{1}[0-9]{1}|0|)[-]{0,1}?[1-9]{1}[0-9]{9}")
            For i = 0 To (mc.Count - 1)
                If Len(TextBox3.Text) = 0 Then
                    TextBox3.Text = mc(i).Value
                Else
                    TextBox3.Text = TextBox3.Text & vbNewLine & mc(i).Value
                End If
            Next

            For i = 0 To (mc1.Count - 1)
                If Len(TextBox3.Text) = 0 Then
                    TextBox3.Text = mc1(i).Value
                Else
                    TextBox3.Text = TextBox3.Text & vbNewLine & mc1(i).Value
                End If
            Next


            txt = TextBox3.Text
            TextBox3.Clear()
            Dim TextLines() As String = txt.Split(vbNewLine).Distinct().ToArray()

            For Each lineTxt In TextLines
                If Len(TextBox3.Text) = 0 Then
                    TextBox3.Text = lineTxt
                Else
                    TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                End If
            Next

        End If
    End Sub


    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Call Refreshfilters()
        TextBox4.Select()
    End Sub


    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Call Refreshfilters()
        TextBox6.Select()
    End Sub


    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Call Refreshfilters()
        TextBox8.Select()
    End Sub


    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If enabledEvents Then
            Call Refreshfilters()
        End If
    End Sub


    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        If enabledEvents Then
            Call Refreshfilters()
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged
        'temp
        If enabledEvents Then
            Call Refreshfilters()
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        If enabledEvents Then
            Call Refreshfilters()
        End If
    End Sub

    'Start-End with Callers****************************************************************
    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        Call Refreshfilters()
        TextBox11.Select()
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        Call Refreshfilters()
        TextBox10.Select()
    End Sub
    'Start-End with Callers****************************************************************


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        txt = TextBox2.Text
        enabledEvents = True
        TextBox2.Select()
    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox15.KeyPress
        If InStr("0,1,2,3,4,5,6,7,8,9", e.KeyChar, CompareMethod.Text) <> 0 Or e.KeyChar = Convert.ToChar(Keys.Back) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        Call Refreshfilters()
        TextBox15.Select()
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If InStr("0,1,2,3,4,5,6,7,8,9", e.KeyChar, CompareMethod.Text) <> 0 Or e.KeyChar = Convert.ToChar(Keys.Back) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Call Refreshfilters()
        TextBox5.Select()
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If (Len(TextBox5.Text) > 0) Then
            Call Refreshfilters()
            TextBox5.Select()
        End If
    End Sub



    Private Sub RadioButton13_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton13.CheckedChanged
        Call Refreshfilters()
    End Sub

    Private Sub RadioButton14_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton14.CheckedChanged
        Call Refreshfilters()
    End Sub

    'Distinct Values
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        txt = TextBox3.Text
        TextBox3.Clear()
        Dim TextLines() As String = txt.Split(vbNewLine).Distinct().ToArray()

        For Each lineTxt In TextLines
            If Len(TextBox3.Text) = 0 Then
                TextBox3.Text = lineTxt
            Else
                TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
            End If
        Next
    End Sub

 

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox3.Text <> String.Empty Then
            Clipboard.SetText(TextBox3.Text)
        Else
            Clipboard.Clear()
        End If

        If Err.Number = 0 Then
            MsgBox("Copied to clipboard")
        Else
            MsgBox("Unable to copy")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim file As System.IO.StreamWriter
        Dim myFileName As String = String.Format("Output_{0}.txt", Now.ToString("MMddyyyy_hhmmss"))
        file = My.Computer.FileSystem.OpenTextFileWriter(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & myFileName, True)
        file.Write(TextBox3.Text)
        file.Close()
        If System.IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & myFileName) = True Then
            Process.Start(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & myFileName)
        Else
            MsgBox("File Does Not Exist")
        End If
    End Sub

    'About Us
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Shell("explorer.exe http://www.megettingerror.blogspot.com/p/about-us.html")
    End Sub

    'Terms of Use
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Shell("explorer.exe http://megettingerror.blogspot.com/p/blog-page_30.html")
    End Sub


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Call Refreshfilters()
    End Sub
End Class
