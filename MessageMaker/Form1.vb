Public Class Form1
    Dim facename As String
    Dim piccount As Integer = 0

    'BETÖLTÉS
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ListBox1.SelectedIndex = 0
    End Sub

    'MUTASS MINDENT
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'NameBox szín
        If ListBox1.SelectedIndex = 0 Then
            TextBox4.Text = "Hermes.name_font_color = ""#00ff1e"""
        ElseIf ListBox1.SelectedIndex = 1 Then
            TextBox4.Text = "Hermes.name_font_color = ""#ffffff"""
        Else
            TextBox4.Text = "Hermes.name_font_color = ""#ff0000"""
        End If


        'Mindenszövegegyben
        TextBox3.Text = ""
        TextBox3.Text = facename & " \n[" & TextBox6.Text & "] " & TextBox2.Text



    End Sub

    'KÖVETKEZŐ KÉP
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If piccount < ListBox2.Items.Count - 1 Then
            piccount += 1
        Else
            piccount = 0
        End If
        ListBox2.SelectedIndex = piccount
        facename = ListBox2.SelectedItem
        facename = facename.Substring(0, facename.Length - facename.LastIndexOf("\") - 5)
        facename = "\f[" & facename & "]"
        PictureBox1.Image = Image.FromFile(TextBox5.Text & ListBox2.SelectedItem)
    End Sub
    Private Function Get_Files(ByVal directory As String, _
                           ByVal recursive As IO.SearchOption, _
                           ByVal ext As String, _
                           ByVal with_word_in_filename As String) As List(Of IO.FileInfo)

        Return IO.Directory.GetFiles(directory, "*" & If(ext.StartsWith("*"), ext.Substring(1), ext), recursive) _
                           .Where(Function(o) o.ToLower.Contains(with_word_in_filename.ToLower)) _
                           .Select(Function(p) New IO.FileInfo(p)).ToList

    End Function

    'MÁS NÉV KERESÉSE
    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        ListBox2.Items.Clear()
        piccount = 0
        For Each file As IO.FileInfo In Get_Files(TextBox5.Text, IO.SearchOption.TopDirectoryOnly, "png", TextBox1.Text)
            ListBox2.Items.Add(file.Name)
        Next
        If ListBox2.Items.Count > 0 Then
            ListBox2.SelectedIndex = 0
            facename = ListBox2.SelectedItem
            facename = facename.Substring(0, facename.Length - facename.LastIndexOf("\") - 5)
            facename = "\f[" & facename & "]"
            PictureBox1.Image = Image.FromFile(TextBox5.Text & ListBox2.SelectedItem)
        End If
        
    End Sub

    'KÉPVÁLTÁS
    Private Sub ListBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        piccount = ListBox2.SelectedIndex
        facename = ListBox2.SelectedItem
        facename = facename.Substring(0, facename.Length - facename.LastIndexOf("\") - 5)
        facename = "\f[" & facename & "]"
        PictureBox1.Image = Image.FromFile(TextBox5.Text & ListBox2.SelectedItem)
    End Sub
End Class
