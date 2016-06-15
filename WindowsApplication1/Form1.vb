Public Class Form1

    Dim txt As String
    Dim txtL1 As String
    Dim txtL2 As String
    Dim txtL3 As String
    Dim txtL4 As String
    Dim enabledEvents As Boolean

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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
            txtL1 = txt
        End If
        enabledEvents = True
    End Sub

    'None
    Private Sub RadioButton11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton11.CheckedChanged
        TextBox3.Text = txt
        txtL1 = TextBox3.Text
    End Sub

    'Tab
    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        TextBox3.Text = Join(Split(txt, vbTab), vbNewLine)
        txtL1 = TextBox3.Text
    End Sub

    'Comma
    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        TextBox3.Text = Join(Split(txt, ","), vbNewLine)
        txtL1 = TextBox3.Text
    End Sub

    'Text
    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        TextBox3.Text = Join(Split(txt, TextBox17.Text), vbNewLine)
        txtL1 = TextBox3.Text
    End Sub

    Private Sub TextBox17_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox17.TextChanged
        If (Len(TextBox17.Text) > 0) Then
            RadioButton4.Enabled = True
            RadioButton11.Select()
        Else
            RadioButton4.Enabled = False
        End If
        TextBox17.Select()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        TextBox3.Clear()
        If (Len(TextBox4.Text) = 0) Then
            TextBox6.Enabled = False
            TextBox3.Text = txt
        Else
            If (Len(TextBox6.Text) <> 0) Then
                If (Len(TextBox8.Text) <> 0) Then
                    If (RadioButton10.Enabled = True) Then
                        Call sixAEight()
                    Else
                        Call sixOEight()
                    End If
                Else
                    If (RadioButton7.Enabled = True) Then
                        Call fourAsix()
                    Else
                        Call fourOsix()
                    End If
                End If
            Else
                Call only4()
            End If

        End If
        TextBox4.Select()
    End Sub


    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        TextBox3.Clear()
        If (Len(TextBox6.Text) = 0) Then
            TextBox8.Enabled = False
            TextBox3.Text = txtL2
        Else
            If (Len(TextBox8.Text) <> 0) Then
                If (RadioButton10.Enabled = True) Then
                    Call sixAEight()
                Else
                    Call sixOEight()
                End If
            Else
                If (RadioButton7.Enabled = True) Then
                    Call fourAsix()
                Else
                    Call fourOsix()
                End If
            End If
        End If
            TextBox6.Select()
    End Sub


    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        TextBox3.Clear()
        If (Len(TextBox8.Text) = 0) Then
            TextBox3.Text = txtL3
        Else
            If (RadioButton10.Enabled = True) Then
                Call sixAeight()
            Else
                Call sixOEight()
            End If
        End If
        TextBox8.Select()
    End Sub

    Sub only4()
        If enabledEvents Then
            Dim TextLines() As String = txtL1.Split(vbNewLine)
            For Each lineTxt In TextLines
                If (InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0) Then
                    TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                End If
            Next
            enabledEvents = False
            RadioButton7.Enabled = True
            RadioButton7.Select()
            RadioButton8.Enabled = True
            TextBox6.Enabled = True
            txtL2 = TextBox3.Text
            enabledEvents = True
        End If
    End Sub

    Sub only6()
        If enabledEvents Then
            Dim TextLines() As String = txtL2.Split(vbNewLine)
            For Each lineTxt In TextLines
                If (InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) Then
                    TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                End If
            Next
            enabledEvents = False
            RadioButton9.Enabled = True
            RadioButton10.Enabled = True
            TextBox8.Enabled = True
            txtL3 = TextBox3.Text
            enabledEvents = True
        End If
    End Sub

    Sub only8()
            Dim TextLines() As String = txtL3.Split(vbNewLine)
            For Each lineTxt In TextLines
                If (InStr(lineTxt, TextBox8.Text, vbTextCompare) > 0) Then
                    TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                End If
            Next
    End Sub


    Sub fourOsix()
        If enabledEvents Then
            TextBox3.Clear()
            If (RadioButton8.Enabled = True) Then
                Dim TextLines() As String = txtL1.Split(vbNewLine)
                For Each lineTxt In TextLines
                    If (InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 Or InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) Then
                        TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                    End If
                Next
                enabledEvents = False
                RadioButton9.Enabled = True
                RadioButton10.Enabled = True
                TextBox8.Enabled = True
                txtL2 = TextBox3.Text
                enabledEvents = True
            End If
        End If
    End Sub

    Sub fourAsix()
        If enabledEvents Then
            TextBox3.Clear()
            If (RadioButton7.Enabled = True) Then
                Dim TextLines() As String = txtL1.Split(vbNewLine)
                For Each lineTxt In TextLines
                    If (InStr(lineTxt, TextBox4.Text, vbTextCompare) > 0 And InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0) Then
                        TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                    End If
                Next
                enabledEvents = False
                RadioButton9.Enabled = True
                RadioButton10.Enabled = True
                TextBox8.Enabled = True
                txtL2 = TextBox3.Text
                enabledEvents = True
            End If
        End If
    End Sub


    Sub sixAEight()
        If enabledEvents Then
            TextBox3.Clear()
            If (RadioButton9.Enabled = True) Then
                Dim TextLines() As String = txtL2.Split(vbNewLine)
                For Each lineTxt In TextLines
                    If (InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0 And InStr(lineTxt, TextBox8.Text, vbTextCompare) > 0) Then
                        TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                    End If
                Next
                txtL3 = TextBox3.Text
            End If
        End If
    End Sub


    Sub sixOEight()
        If enabledEvents Then
            TextBox3.Clear()
            If (RadioButton9.Enabled = True) Then
                Dim TextLines() As String = txtL2.Split(vbNewLine)
                For Each lineTxt In TextLines
                    If (InStr(lineTxt, TextBox6.Text, vbTextCompare) > 0 Or InStr(lineTxt, TextBox8.Text, vbTextCompare) > 0) Then
                        TextBox3.Text = TextBox3.Text & vbNewLine & lineTxt
                    End If
                Next
                txtL3 = TextBox3.Text
            End If
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        If enabledEvents Then
            Call fourAsix()
        End If
    End Sub


    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged
        If enabledEvents Then
            Call fourOsix()
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton10.CheckedChanged
        'temp
        If enabledEvents Then
            Call sixAEight()
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton9.CheckedChanged
        If enabledEvents Then
            Call sixOEight()
        End If
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        TextBox3.Clear()
        If (Len(TextBox11.Text) = 0) Then
            TextBox3.Text = txtL3
        Else
            If (RadioButton10.Enabled = True) Then
                Call sixAEight()
            Else
                Call sixOEight()
            End If
        End If
        TextBox11.Select()
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub
End Class
