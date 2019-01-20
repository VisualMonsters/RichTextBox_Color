Public Class Form1

    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByRef lParam As Point) As Integer
    Const WM_USER As Integer = &H400
    Const EM_GETSCROLLPOS As Integer = WM_USER + 221
    Const EM_SETSCROLLPOS As Integer = WM_USER + 222

    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Integer) As Integer

    Dim specialSigns() As String = {"+", "-", "*", "/", "%", "=", "test"}
    'obsługuje zdarzenie wklejenia tekstu
    Dim paste As Boolean = False
    Dim pasteStartPosition As Integer
    Dim lineCount As Integer

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        Dim RTB1SP As Point
        SendMessage(RichTextBox1.Handle, EM_GETSCROLLPOS, 0, RTB1SP)
        LockWindowUpdate(RichTextBox1.Handle.ToInt32) 'blokuje
        If paste Then
            For i As Integer = 0 To lineCount - 1
                HighLightOperatorKey(RichTextBox1, pasteStartPosition)
                pasteStartPosition += 1
            Next
            paste = False
        Else
            If RichTextBox1.TextLength > 0 Then
                HighLightOperatorKey(RichTextBox1)
            End If
        End If
        LockWindowUpdate(0) 'odblokuje
        SendMessage(RichTextBox1.Handle, EM_SETSCROLLPOS, 0, RTB1SP)
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown

        If e.Modifiers = Keys.Control AndAlso e.KeyCode = Keys.V Then
            Dim lines As String() = My.Computer.Clipboard.GetText().Split(Environment.NewLine)
            pasteStartPosition = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart)
            lineCount = lines.Count
            paste = True
        End If
    End Sub

#Region "Kolorowanie znaków specjalnych"
    Public Sub HighLightOperatorKey(ByVal rtb As RichTextBox, Optional ByVal paste As Integer = -1)
        'pobiera aktualną pozycje kursora

        Dim startPosition As Integer = rtb.SelectionStart

        rtb.Select(startPosition - 1, 1)
        rtb.SelectionColor = Color.White

        'sprawdza zawartość okna tekstowego w poszukiwaniu elementów tablicy
        For Each sign As String In specialSigns
            Dim position As Integer
            If paste = -1 Then
                position = rtb.GetFirstCharIndexFromLine(rtb.GetLineFromCharIndex(startPosition))
            Else
                position = rtb.GetFirstCharIndexFromLine(paste)
            End If

            'pobiera numer lini                                 ==== RichTextBox1.GetLineFromCharIndex(startPosition)
            'pobiera numer pierwszego znaku w liniii            ==== RichTextBox1.GetFirstCharIndexOfCurrentLine()
            'pobiera numer pierwszego znaku w wybranej linni    ==== RichTextBox1.GetFirstCharIndexFromLine(numer)
            Do While rtb.Text.IndexOf(sign, position) >= 0
                position = rtb.Text.IndexOf(sign, position)
                rtb.Select(position + 1, 0)
                'nasz tekst jest dodatkowo sprawdzani 
                'a zaznaczenie jest omijane gdy znak ma już kolor
                If Not rtb.SelectionColor = Color.Red Then
                    rtb.Select(position, sign.Length)
                    'rtb.SelectionFont = New Font("Courier New", 12, FontStyle.Regular)
                    rtb.SelectionColor = Color.Red  'kolor czerwony dla znaku
                End If
                position += 1
            Loop
        Next

        rtb.SelectionLength = 0
        rtb.SelectionStart = startPosition
        'reset koloru
        rtb.SelectionColor = Color.White
    End Sub


#End Region

End Class
