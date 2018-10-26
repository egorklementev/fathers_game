Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class Form1
    Dim Pole(10, 10) As Short
    Dim PoleTmp(10, 10) As Short
    Dim TopPole(10) As Short
    Dim ColTmp As Short
    Dim ColDel As Short
    Dim Count As Single
    Dim Bonus As Single
    Dim GameOver As Boolean = False
    Dim m_x, m_y As Integer
    Dim x1 As Short
    Dim x_EndGame, y_EndGame As Short
    Dim iDD As Integer
    Dim imgRED = Image.FromFile("red.jpg")
    Dim imgBLUE = Image.FromFile("blue.jpg")
    Dim imgGREEN = Image.FromFile("green.jpg")
    Dim imgBLACK = Image.FromFile("black.jpg")
    Dim imgYELLOW = Image.FromFile("yellow.jpg")
    Dim imgVIOLET = Image.FromFile("violet.jpg")
    Dim imgBROWN = Image.FromFile("brown.jpg")
    Dim imgBOMB = Image.FromFile("bomb.jpg")
    Dim imgGORIZONT = Image.FromFile("gor.jpg")
    Dim imgVERTIKAL = Image.FromFile("vert.jpg")
    Dim imgKREST = Image.FromFile("gor_vert.jpg")
    Dim imgBombRED = Image.FromFile("red_bomb.jpg")
    Dim imgBombGREEN = Image.FromFile("green_bomb.jpg")
    Dim imgBombBLUE = Image.FromFile("blue_bomb.jpg")
    Dim imgBombYELLOW = Image.FromFile("yellow_bomb.jpg")
    Dim imgBombVIOLET = Image.FromFile("violet_bomb.jpg")
    Dim imgBombBrown = Image.FromFile("brown_bomb.jpg")

    Dim imgGameOver = Image.FromFile("gameover.jpg")

    Dim imgtopRED = Image.FromFile("top_red.jpg")
    Dim imgtopBLUE = Image.FromFile("top_blue.jpg")
    Dim imgtopGREEN = Image.FromFile("top_green.jpg")
    Dim imgtopBLACK = Image.FromFile("top_black.jpg")
    Dim imgtopYELLOW = Image.FromFile("top_yellow.jpg")
    Dim imgtopVIOLET = Image.FromFile("top_violet.jpg")
    Dim imgtopBROWN = Image.FromFile("top_brown.jpg")
    Dim imgtopSUR = Image.FromFile("top_sur.jpg")

    Dim level As Short = 3

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        For i = 1 To Int(Rnd() * TimeOfDay.Second)
            x1 = Rnd()
        Next
        For x = 1 To 10
            TopPole(x) = 0
            For y = 1 To 10
                Pole(x, y) = 0
            Next
        Next
        x1 = 1
    End Sub

    Private Sub Del_Left()
        If m_x = 1 Then Exit Sub
        m_x -= 1
        If Pole(m_x, m_y) = ColTmp Then
            ColDel += 1
            Pole(m_x, m_y) = 0
            Del_Left()
            Del_Up()
            Del_Right()
            Del_Down()
        End If
        m_x += 1
    End Sub
    Private Sub Del_Right()
        If m_x = 10 Then Exit Sub
        m_x += 1
        If Pole(m_x, m_y) = ColTmp Then
            ColDel += 1
            Pole(m_x, m_y) = 0
            Del_Right()
            Del_Down()
            Del_Left()
            Del_Up()
        End If
        m_x -= 1
    End Sub
    Private Sub Del_Up()
        If m_y = 1 Then Exit Sub
        m_y -= 1
        If Pole(m_x, m_y) = ColTmp Then
            ColDel += 1
            Pole(m_x, m_y) = 0
            Del_Up()
            Del_Right()
            Del_Down()
            Del_Left()
        End If
        m_y += 1
    End Sub
    Private Sub Del_Down()
        If m_y = 10 Then Exit Sub
        m_y += 1
        If Pole(m_x, m_y) = ColTmp Then
            ColDel += 1
            Pole(m_x, m_y) = 0
            Del_Down()
            Del_Left()
            Del_Up()
            Del_Right()
        End If
        m_y -= 1
    End Sub
    Private Sub ZipLeft()
        Timer2.Enabled = False
        Timer1.Enabled = False
        Dim y1, x1, c As Short
        Dim test As Boolean = True
        For c = 1 To 4
            For x1 = 5 To 2 Step -1
                For y1 = 10 To 1 Step -1
                    If Pole(x1, y1) <> 0 Then
                        test = False
                        Exit For
                    End If
                Next
                If test = True Then
                    For y1 = 1 To 10
                        Pole(x1, y1) = Pole(x1 - 1, y1)
                        Pole(x1 - 1, y1) = 0
                    Next
                End If
                test = True
            Next
        Next
        Timer2.Enabled = True
        Timer1.Enabled = True
    End Sub
    Private Sub ZipRight()
        Timer2.Enabled = False
        Timer1.Enabled = False
        Dim y1, x1, c As Short
        Dim test As Boolean = True
        For c = 1 To 4
            For x1 = 6 To 9
                For y1 = 10 To 1 Step -1
                    If Pole(x1, y1) <> 0 Then
                        test = False
                        Exit For
                    End If
                Next
                If test = True Then
                    For y1 = 1 To 10
                        Pole(x1, y1) = Pole(x1 + 1, y1)
                        Pole(x1 + 1, y1) = 0
                    Next
                End If
                test = True
            Next
        Next
        Timer2.Enabled = True
        Timer1.Enabled = True
    End Sub
    Private Sub EndGame()
        Timer1.Enabled = False
        Timer2.Enabled = False
        For x = 1 To 10
            TopPole(x) = 0
        Next
        PictureBox2.Refresh()
        Timer3.Enabled = True
    End Sub
    Private Sub Bomb()
        Pole(m_x, m_y) = 0
        If m_x < 10 Then
            Pole(m_x + 1, m_y) = 0
            ColDel += 1
        End If
        If m_x < 10 And m_y < 10 Then
            Pole(m_x + 1, m_y + 1) = 0
            ColDel += 1
        End If
        If m_y < 10 Then
            Pole(m_x, m_y + 1) = 0
            ColDel += 1
        End If
        If m_x > 1 And m_y < 10 Then
            Pole(m_x - 1, m_y + 1) = 0
            ColDel += 1
        End If
        If m_x > 1 Then
            Pole(m_x - 1, m_y) = 0
            ColDel += 1
        End If
        If m_x > 1 And m_y > 1 Then
            Pole(m_x - 1, m_y - 1) = 0
            ColDel += 1
        End If
        If m_y > 1 Then
            Pole(m_x, m_y - 1) = 0
            ColDel += 1
        End If
        If m_x < 10 And m_y > 1 Then
            Pole(m_x + 1, m_y - 1) = 0
            ColDel += 1
        End If
    End Sub
    Private Sub Gorizont()
        For i = 1 To 10
            If Pole(i, m_y) <> 0 Then ColDel += 1
            Pole(i, m_y) = 0
        Next
    End Sub
    Private Sub Vertical()
        For i = 1 To 10
            If Pole(m_x, i) <> 0 Then ColDel += 1
            Pole(m_x, i) = 0
        Next
    End Sub
    Private Sub Krest()
        For i = 1 To 10
            If Pole(i, m_y) <> 0 Then ColDel += 1
            If Pole(m_x, i) <> 0 Then ColDel += 1
            Pole(i, m_y) = 0
            Pole(m_x, i) = 0
        Next
    End Sub
    Private Sub ColorDelete()
        ColTmp = Pole(m_x, m_y) - 14
        Pole(m_x, m_y) = 0
        ColDel += 1
        For x = 1 To 10
            For y = 1 To 10
                If Pole(x, y) = ColTmp Then
                    Pole(x, y) = 0
                    ColDel += 1
                End If

            Next
        Next
    End Sub
    Private Sub DropDown()
        iDD = 1
        Timer2.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ' Верхняя полоска
        Dim c As Short
        TopPole(x1) = Int(Rnd() * level) + 1

        If level = 4 Then
            c = Int(Rnd() * 50) + 1
            If c = 11 Then TopPole(x1) = 11

            If c = 15 Then TopPole(x1) = 15
            If c = 16 Then TopPole(x1) = 16
            If c = 17 Then TopPole(x1) = 17
            If c = 18 Then TopPole(x1) = 18
        End If
        If level = 5 Then
            c = Int(Rnd() * 50) + 1
            If c = 11 Then TopPole(x1) = 11
            If c = 12 Then TopPole(x1) = 12
            If c = 13 Then TopPole(x1) = 13

            If c = 15 Then TopPole(x1) = 15
            If c = 16 Then TopPole(x1) = 16
            If c = 17 Then TopPole(x1) = 17
            If c = 18 Then TopPole(x1) = 18
            If c = 19 Then TopPole(x1) = 19
        End If
        If level = 6 Then
            c = Int(Rnd() * 50) + 1
            If c = 11 Then TopPole(x1) = 11
            If c = 12 Then TopPole(x1) = 12
            If c = 13 Then TopPole(x1) = 13
            If c = 14 Then TopPole(x1) = 14
        End If
        If level = 7 Then
            c = Int(Rnd() * 50) + 1
            If c = 11 Then TopPole(x1) = 11
            If c = 12 Then TopPole(x1) = 12
            If c = 13 Then TopPole(x1) = 13
            If c = 14 Then TopPole(x1) = 14
        End If
        Me.PictureBox2.Refresh()
        x1 += 1
        If x1 = 11 Then
            Me.Timer1.Interval = 500
            x1 = 1
            Dim x2 As Short
            For x2 = 1 To 10
                If Pole(x2, 1) <> 0 Then
                    EndGame()
                    Exit Sub
                End If
                Pole(x2, 1) = TopPole(x2)
                TopPole(x2) = 0
            Next
            Me.PictureBox1.Refresh()
            DropDown()
        End If
    End Sub
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        ' ПАДЕНИЕ
        For y = 10 To 2 Step -1
            For x = 1 To 10
                If Pole(x, y) = 0 Then
                    Pole(x, y) = Pole(x, y - 1)
                    Pole(x, y - 1) = 0
                End If
            Next
        Next
        iDD = +1
        If iDD = 10 Then
            Me.Timer2.Enabled = False
        End If
        Me.PictureBox1.Refresh()
    End Sub
    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        y_EndGame += 1
        For x_EndGame = 1 To 10
            Pole(x_EndGame, y_EndGame) = 20
            PictureBox1.Refresh()
        Next
        If y_EndGame = 10 Then
            Timer3.Enabled = False
            y_EndGame = 0
            GameOver = True
            Me.PictureBox1.Refresh()
        End If
    End Sub

    Private Sub PictureBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseClick
        Array.Copy(Pole, PoleTmp, Pole.Length)
        ColDel = 0
        Bonus = 0
        m_x = Int(e.X / 50) + 1
        m_y = Int(e.Y / 50) + 1
        If Pole(m_x, m_y) = 0 Then Exit Sub

        If Pole(m_x, m_y) = 11 Then
            Bomb()
            DropDown()
            ZipLeft()
            ZipRight()
            Count = Count + ColDel + Bonus
            Me.PictureBox1.Refresh()
            Exit Sub
        End If

        If Pole(m_x, m_y) = 12 Then
            Gorizont()
            DropDown()
            Count = Count + ColDel + Bonus
             Me.PictureBox1.Refresh()
            Exit Sub
        End If

        If Pole(m_x, m_y) = 13 Then
            Vertical()
            ZipLeft()
            ZipRight()
            Count = Count + ColDel + Bonus
            Me.PictureBox1.Refresh()
            Exit Sub
        End If

        If Pole(m_x, m_y) = 14 Then
            Krest()
            DropDown()
            ZipLeft()
            ZipRight()
            Count = Count + ColDel + Bonus
            Me.PictureBox1.Refresh()
            Exit Sub
        End If

        If Pole(m_x, m_y) = 15 Or Pole(m_x, m_y) = 16 Or Pole(m_x, m_y) = 17 Or Pole(m_x, m_y) = 18 Or Pole(m_x, m_y) = 19 Then
            ColorDelete()
            DropDown()
            ZipLeft()
            ZipRight()
            Count = Count + ColDel + Bonus
            Me.PictureBox1.Refresh()
            Exit Sub
        End If

        If (m_x > 0 And m_x < 11) Or (m_y > 0 And m_y < 11) Then
            ColTmp = Pole(m_x, m_y)
            Pole(m_x, m_y) = 0
            ColDel = 1
        End If
        Del_Left()
        Del_Right()
        Del_Up()
        Del_Down()
        If ColDel > 2 Then
            Me.PictureBox1.Refresh()
            DropDown()
            Me.PictureBox1.Refresh()
            ZipLeft()
            ZipRight()

            Select Case ColDel
                Case 20 To 30
                    Bonus = 50
                Case 31 To 40
                    Bonus = 100
                Case 41 To 50
                    Bonus = 150
                Case 51 To 60
                    Bonus = 200
                Case 61 To 70
                    Bonus = 250
                Case 71 To 80
                    Bonus = 300
                Case 81 To 90
                    Bonus = 350
                Case 91 To 100
                    Bonus = 400
            End Select
            Count = Count + ColDel + Bonus
            Select Case Count
                Case 300 To 599
                    level = 4
                Case 600 To 899
                    level = 5
                Case 900 To 1199
                    level = 6
                Case 1200 To 1499
                    level = 7
                Case 1500 To 1990
                    level = 8
            End Select
            Me.PictureBox1.Refresh()
        Else
            Array.Copy(PoleTmp, Pole, Pole.Length)
        End If
    End Sub
    Private Sub PictureBox1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint
        Dim g As Graphics = e.Graphics
        'g.SmoothingMode = SmoothingMode.HighQuality
        Dim MainFont As New Font("Arial", 40, Font.Bold, System.Drawing.GraphicsUnit.Point)
        Dim Main1Font As New Font("Arial", 10, Font.Bold, System.Drawing.GraphicsUnit.Point)
        Dim Main2Font As New Font("Times New Roman", 100, Font.Bold, System.Drawing.GraphicsUnit.Point)
        For x = 1 To 10
            For y = 1 To 10
                Select Case Pole(x, y)
                    Case 0
                        g.DrawImage(imgBLACK, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 1
                        g.DrawImage(imgRED, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 2
                        g.DrawImage(imgGREEN, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 3
                        g.DrawImage(imgBLUE, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 4
                        g.DrawImage(imgYELLOW, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 5
                        g.DrawImage(imgVIOLET, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 6
                        g.DrawImage(imgBROWN, (x * 50) - 50, (y * 50) - 50, 49, 49)

                    Case 11
                        g.DrawImage(imgBOMB, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 12
                        g.DrawImage(imgGORIZONT, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 13
                        g.DrawImage(imgVERTIKAL, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 14
                        g.DrawImage(imgKREST, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 15
                        g.DrawImage(imgBombRED, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 16
                        g.DrawImage(imgBombGREEN, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 17
                        g.DrawImage(imgBombBLUE, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 18
                        g.DrawImage(imgBombYELLOW, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 19
                        g.DrawImage(imgBombVIOLET, (x * 50) - 50, (y * 50) - 50, 49, 49)
                    Case 20
                        g.DrawImage(imgGameOver, (x * 50) - 50, (y * 50) - 50, 49, 49)
                End Select
            Next
        Next
        If Bonus <> 0 Then
            Dim stringSize As New SizeF
            stringSize = g.MeasureString(CStr("BONUS: " & Bonus), MainFont)
            g.DrawString("BONUS: " & Bonus, MainFont, Brushes.White, (Me.PictureBox1.Width - stringSize.Width) / 2, 200)
        End If
        Me.Label2.Text = Count
        Me.Label3.Text = level - 2
        If GameOver = True Then
            Dim stringSize As New SizeF
            stringSize = g.MeasureString("GAME", Main2Font)
            g.DrawString("GAME", Main2Font, Brushes.Black, ((Me.PictureBox1.Width - stringSize.Width) / 2) + 10, 100 + 10)
            g.DrawString("GAME", Main2Font, Brushes.Red, (Me.PictureBox1.Width - stringSize.Width) / 2, 100)
            stringSize = g.MeasureString("OVER", Main2Font)
            g.DrawString("OVER", Main2Font, Brushes.Black, ((Me.PictureBox1.Width - stringSize.Width) / 2) + 10, 250 + 10)
            g.DrawString("OVER", Main2Font, Brushes.Red, (Me.PictureBox1.Width - stringSize.Width) / 2, 250)
            GameOver = False
        End If
    End Sub

    Private Sub PictureBox2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseClick
        Me.Timer1.Interval = 50
    End Sub
    Private Sub PictureBox2_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox2.Paint
        Dim g As Graphics = e.Graphics
        For xPB2 = 1 To 10
            'Dim br2 As New SolidBrush(TopPole(xPB2))
            'g.FillRectangle(br2, (xPB2 * 50) - 50, 0, 49, 20)
            Select Case TopPole(xPB2)
                Case 0
                    g.DrawImage(imgtopBLACK, (xPB2 * 50) - 50, 0, 49, 20)
                Case 1
                    g.DrawImage(imgtopRED, (xPB2 * 50) - 50, 0, 49, 20)
                Case 2
                    g.DrawImage(imgtopGREEN, (xPB2 * 50) - 50, 0, 49, 20)
                Case 3
                    g.DrawImage(imgtopBLUE, (xPB2 * 50) - 50, 0, 49, 20)
                Case 4
                    g.DrawImage(imgtopYELLOW, (xPB2 * 50) - 50, 0, 49, 20)
                Case 5
                    g.DrawImage(imgtopVIOLET, (xPB2 * 50) - 50, 0, 49, 20)
                Case 6
                    g.DrawImage(imgtopBROWN, (xPB2 * 50) - 50, 0, 49, 20)
                Case 11 To 19
                    g.DrawImage(imgtopSUR, (xPB2 * 50) - 50, 0, 49, 20)
                Case Else
                    g.DrawImage(imgtopSUR, (xPB2 * 50) - 50, 0, 49, 20)
            End Select
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        For x = 1 To 10
            TopPole(x) = 0
            For y = 1 To 10
                Pole(x, y) = 0
            Next
        Next
        x1 = 1
        ColDel = 0
        Count = 0
        Bonus = 0
        level = 3
        Timer1.Enabled = True
        Timer2.Enabled = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Timer1.Enabled = False Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim g As Graphics = e.Graphics
        Dim p As New Pen(Brushes.Yellow, 10)
        g.DrawRectangle(p, Me.PictureBox1.Left, Me.PictureBox1.Top, 500, 500)
        g.DrawRectangle(p, Me.PictureBox2.Left, Me.PictureBox2.Top, 500, 20)
    End Sub
End Class
