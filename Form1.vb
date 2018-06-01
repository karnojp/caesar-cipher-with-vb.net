Public Class Form1

    Function EncryptDecrypt(ByVal text1 As String, ByVal key As String, ByVal isEncrypt As Boolean) As String
        Dim char1 As String
        Dim char2 As String
        Dim cKey As Byte
        Dim strLength As Integer
        Dim Result As String = ""
        Dim x As Integer = -1
        If text1 <> "" And IsNumeric(key) Then

            strLength = text1.Length
            For i As Integer = 0 To strLength - 1
                char1 = text1.Substring(i, 1)
                If x < key.Length - 1 Then
                    x = x + 1
                Else
                    x = 0
                End If
                cKey = Val(key.Substring(x, 1))
                If isEncrypt Then
                    char2 = Chr(Asc(char1) + cKey)
                Else
                    char2 = Chr(Asc(char1) - cKey)
                End If
                Result &= char2
            Next
        Else
            MsgBox("Enter text or key!")
        End If
        Return Result
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox3 .Text = EncryptDecrypt(TextBox1 .Text, TextBox2.Text, True)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox3.Text = EncryptDecrypt(TextBox1.Text, TextBox2.Text, False)
    End Sub
End Class
