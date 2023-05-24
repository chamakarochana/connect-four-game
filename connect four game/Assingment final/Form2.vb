Imports System.IO
Imports System.Diagnostics.Metrics
Public Class Form2
    'declaring the arrays
    Dim c1(7) As String
    Dim c2(7) As String
    Dim c3(7) As String
    Dim c4(7) As String
    Dim c5(7) As String
    Dim c6(7) As String
    Dim c7(7) As String

    Dim countAvert As Integer
    Dim countBvert As Integer
    Dim countAhorizontal As Integer
    Dim countBhorizontal As Integer



    Dim winVertical(8) As String
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        'changing the colours of the labels to green
        Label1.ForeColor = Color.Green
        Label2.ForeColor = Color.Green
        Label3.ForeColor = Color.Green
        Label4.ForeColor = Color.Green
        Label5.ForeColor = Color.Green
        Label6.ForeColor = Color.Green
        Label7.ForeColor = Color.Green

        'condition to check the selection of the combobox
        If ComboBox1.SelectedIndex = 0 Then
            'changing the label color to red 
            Label1.ForeColor = Color.Red
            'for loop to store number 1 in the array
            For i = 0 To 6

                If c1(i) = "" Then

                    c1(i) = "1"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 1 Then
            'changing the label color to red 
            Label2.ForeColor = Color.Red
            'for loop to store number 1 in the array
            For ii = 0 To 6

                If c2(ii) = "" Then

                    c2(ii) = "1"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 2 Then
            'changing the label color to red 
            Label3.ForeColor = Color.Red
            'for loop to store number 1 in the array
            For iii = 0 To 6

                If c3(iii) = "" Then

                    c3(iii) = "1"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 3 Then
            'changing the label color to red 
            Label4.ForeColor = Color.Red
            'for loop to store number 1 in the array
            For iv = 0 To 6

                If c4(iv) = "" Then

                    c4(iv) = "1"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 4 Then
            'changing the label color to red 
            Label5.ForeColor = Color.Red
            'for loop to store number 1 in the array
            For v = 0 To 6

                If c5(v) = "" Then

                    c5(v) = "1"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 5 Then
            'changing the label color to red 
            Label6.ForeColor = Color.Red
            'for loop to store number 1 in the array
            For vi = 0 To 6

                If c6(vi) = "" Then

                    c6(vi) = "1"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 6 Then
            'changing the label color to red 
            Label7.ForeColor = Color.Red
            'for loop to store number 1 in the array
            For vii = 0 To 6

                If c7(vii) = "" Then

                    c7(vii) = "1"

                    Exit For

                End If

            Next
        End If

        'assigning the number to the text box
        TextBox7.Text = c1(0)
        TextBox6.Text = c1(1)
        TextBox5.Text = c1(2)
        TextBox4.Text = c1(3)
        TextBox3.Text = c1(4)
        TextBox2.Text = c1(5)
        TextBox1.Text = c1(6)


        'assigning the number to the text box
        TextBox8.Text = c2(0)
        TextBox9.Text = c2(1)
        TextBox10.Text = c2(2)
        TextBox11.Text = c2(3)
        TextBox12.Text = c2(4)
        TextBox13.Text = c2(5)
        TextBox14.Text = c2(6)


        'assigning the number to the text box
        TextBox15.Text = c3(0)
        TextBox16.Text = c3(1)
        TextBox17.Text = c3(2)
        TextBox18.Text = c3(3)
        TextBox19.Text = c3(4)
        TextBox20.Text = c3(5)
        TextBox21.Text = c3(6)


        'assigning the number to the text box
        TextBox22.Text = c4(0)
        TextBox23.Text = c4(1)
        TextBox24.Text = c4(2)
        TextBox25.Text = c4(3)
        TextBox26.Text = c4(4)
        TextBox27.Text = c4(5)
        TextBox28.Text = c4(6)


        'assigning the number to the text box
        TextBox29.Text = c5(0)
        TextBox30.Text = c5(1)
        TextBox31.Text = c5(2)
        TextBox32.Text = c5(3)
        TextBox33.Text = c5(4)
        TextBox34.Text = c5(5)
        TextBox35.Text = c5(6)


        'assigning the number to the text box
        TextBox36.Text = c6(0)
        TextBox37.Text = c6(1)
        TextBox38.Text = c6(2)
        TextBox39.Text = c6(3)
        TextBox40.Text = c6(4)
        TextBox41.Text = c6(5)
        TextBox42.Text = c6(6)


        'assigning the number to the text box
        TextBox43.Text = c7(0)
        TextBox44.Text = c7(1)
        TextBox45.Text = c7(2)
        TextBox46.Text = c7(3)
        TextBox47.Text = c7(4)
        TextBox48.Text = c7(5)
        TextBox49.Text = c7(6)



        countAvert = 0
        countBvert = 0
        countAhorizontal = 0
        countBhorizontal = 0



        'checking the winning combination horizontally 
        For horizontal As Integer = 0 To 4
            If c1(horizontal) = "1" And c2(horizontal) = "1" And c3(horizontal) = "1" And c4(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c1(horizontal) = "2" And c2(horizontal) = "2" And c3(horizontal) = "2" And c4(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10
            ElseIf c2(horizontal) = "1" And c3(horizontal) = "1" And c4(horizontal) = "1" And c5(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c2(horizontal) = "2" And c3(horizontal) = "2" And c4(horizontal) = "2" And c5(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10
            ElseIf c3(horizontal) = "1" And c4(horizontal) = "1" And c5(horizontal) = "1" And c6(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c3(horizontal) = "2" And c4(horizontal) = "2" And c5(horizontal) = "2" And c6(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10
            ElseIf c4(horizontal) = "1" And c5(horizontal) = "1" And c6(horizontal) = "1" And c7(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c4(horizontal) = "2" And c5(horizontal) = "2" And c6(horizontal) = "2" And c7(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10

            End If

        Next

        'checking the winning combination vertically 
        For vert As Integer = 0 To 7
            If c1(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c1(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c2(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c2(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c3(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c3(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c4(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c4(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c5(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c5(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c6(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c6(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c7(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c7(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            End If
        Next

        If countAvert > 3 Then
            MsgBox(" Congratulations Player 1 wins ")
        ElseIf countBvert > 3 Then
            MsgBox(" Congratulations Player 2 wins ")
        End If

        If countAhorizontal > 3 Then
            MsgBox(" Congratulations Player 1 wins ")
        ElseIf countBhorizontal > 3 Then
            MsgBox(" Congratulations Player 2 wins ")
        End If





    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'changing the colours of the labels to green
        Label1.ForeColor = Color.Green
        Label2.ForeColor = Color.Green
        Label3.ForeColor = Color.Green
        Label4.ForeColor = Color.Green
        Label5.ForeColor = Color.Green
        Label6.ForeColor = Color.Green
        Label7.ForeColor = Color.Green

        'condition to check the selection of the combobox
        If ComboBox1.SelectedIndex = 0 Then
            'changing the label color to red
            Label1.ForeColor = Color.Red
            'for loop to store number 2 in the array
            For i = 0 To 6

                If c1(i) = "" Then

                    c1(i) = "2"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 1 Then
            'changing the label color to red
            Label2.ForeColor = Color.Red
            'for loop to store number 2 in the array
            For ii = 0 To 6

                If c2(ii) = "" Then

                    c2(ii) = "2"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 2 Then
            'changing the label color to red
            Label3.ForeColor = Color.Red
            'for loop to store number 2 in the array
            For iii = 0 To 6

                If c3(iii) = "" Then

                    c3(iii) = "2"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 3 Then
            'changing the label color to red
            Label4.ForeColor = Color.Red
            'for loop to store number 2 in the array
            For iv = 0 To 6

                If c4(iv) = "" Then

                    c4(iv) = "2"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 4 Then
            'changing the label color to red
            Label5.ForeColor = Color.Red
            'for loop to store number 2 in the array
            For v = 0 To 6

                If c5(v) = "" Then

                    c5(v) = "2"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 5 Then
            'changing the label color to red
            Label6.ForeColor = Color.Red
            'for loop to store number 2 in the array
            For vi = 0 To 6
                System.Console.WriteLine(c6(vi))

                If c6(vi) = "" Then

                    c6(vi) = "2"

                    Exit For

                End If

            Next

            'condition to check the selection of the combobox
        ElseIf ComboBox1.SelectedIndex = 6 Then
            'changing the label color to red
            Label7.ForeColor = Color.Red
            'for loop to store number 2 in the array
            For vii = 0 To 6

                If c7(vii) = "" Then

                    c7(vii) = "2"

                    Exit For
                End If
            Next

        End If


        'assigning the number to the text box
        TextBox7.Text = c1(0)
        TextBox6.Text = c1(1)
        TextBox5.Text = c1(2)
        TextBox4.Text = c1(3)
        TextBox3.Text = c1(4)
        TextBox2.Text = c1(5)
        TextBox1.Text = c1(6)


        'assigning the number to the text box
        TextBox8.Text = c2(0)
        TextBox9.Text = c2(1)
        TextBox10.Text = c2(2)
        TextBox11.Text = c2(3)
        TextBox12.Text = c2(4)
        TextBox13.Text = c2(5)
        TextBox14.Text = c2(6)


        'assigning the number to the text box
        TextBox15.Text = c3(0)
        TextBox16.Text = c3(1)
        TextBox17.Text = c3(2)
        TextBox18.Text = c3(3)
        TextBox19.Text = c3(4)
        TextBox20.Text = c3(5)
        TextBox21.Text = c3(6)


        'assigning the number to the text box
        TextBox22.Text = c4(0)
        TextBox23.Text = c4(1)
        TextBox24.Text = c4(2)
        TextBox25.Text = c4(3)
        TextBox26.Text = c4(4)
        TextBox27.Text = c4(5)
        TextBox28.Text = c4(6)


        'assigning the number to the text box
        TextBox29.Text = c5(0)
        TextBox30.Text = c5(1)
        TextBox31.Text = c5(2)
        TextBox32.Text = c5(3)
        TextBox33.Text = c5(4)
        TextBox34.Text = c5(5)
        TextBox35.Text = c5(6)


        'assigning the number to the text box
        TextBox36.Text = c6(0)
        TextBox37.Text = c6(1)
        TextBox38.Text = c6(2)
        TextBox39.Text = c6(3)
        TextBox40.Text = c6(4)
        TextBox41.Text = c6(5)
        TextBox42.Text = c6(6)


        'assigning the number to the text box
        TextBox43.Text = c7(0)
        TextBox44.Text = c7(1)
        TextBox45.Text = c7(2)
        TextBox46.Text = c7(3)
        TextBox47.Text = c7(4)
        TextBox48.Text = c7(5)
        TextBox49.Text = c7(6)

        countAvert = 0
        countBvert = 0
        countAhorizontal = 0
        countBhorizontal = 0



        'checking the winning combination horizontally 
        For horizontal As Integer = 0 To 4
            If c1(horizontal) = "1" And c2(horizontal) = "1" And c3(horizontal) = "1" And c4(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c1(horizontal) = "2" And c2(horizontal) = "2" And c3(horizontal) = "2" And c4(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10
            ElseIf c2(horizontal) = "1" And c3(horizontal) = "1" And c4(horizontal) = "1" And c5(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c2(horizontal) = "2" And c3(horizontal) = "2" And c4(horizontal) = "2" And c5(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10
            ElseIf c3(horizontal) = "1" And c4(horizontal) = "1" And c5(horizontal) = "1" And c6(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c3(horizontal) = "2" And c4(horizontal) = "2" And c5(horizontal) = "2" And c6(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10
            ElseIf c4(horizontal) = "1" And c5(horizontal) = "1" And c6(horizontal) = "1" And c7(horizontal) = "1" Then
                countBhorizontal = 0
                countAhorizontal = 10
            ElseIf c4(horizontal) = "2" And c5(horizontal) = "2" And c6(horizontal) = "2" And c7(horizontal) = "2" Then
                countAhorizontal = 0
                countBhorizontal = 10

            End If

        Next

        'checking the winning combination vertically 
        For vert As Integer = 0 To 7
            If c1(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c1(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c2(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c2(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c3(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c3(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c4(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c4(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c5(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c5(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c6(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c6(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            ElseIf c7(vert) = "1" Then
                countBvert = 0
                countAvert = countAvert + 1
            ElseIf c7(vert) = "2" Then
                countAvert = 0
                countBvert = countBvert + 1
            End If
        Next

        If countAvert > 3 Then
            MsgBox(" Congratulations Player 1 wins ")
        ElseIf countBvert > 3 Then
            MsgBox(" Congratulations Player 2 wins ")
        End If

        If countAhorizontal > 3 Then
            MsgBox(" Congratulations Player 1 wins ")
        ElseIf countBhorizontal > 3 Then
            MsgBox(" Congratulations Player 2 wins ")
        End If







    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'saving the game data 
        System.IO.File.WriteAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c1.txt", c1)
        System.IO.File.WriteAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c2.txt", c2)
        System.IO.File.WriteAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c3.txt", c3)
        System.IO.File.WriteAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c4.txt", c4)
        System.IO.File.WriteAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c5.txt", c5)
        System.IO.File.WriteAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c6.txt", c6)
        System.IO.File.WriteAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c7.txt", c7)



    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        'loading the game
        Try

            'reading all the lines in the file 
            c1 = File.ReadAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c1.txt")
            'for loop to load the game 
            TextBox7.Text = c1(0)
            TextBox6.Text = c1(1)
            TextBox5.Text = c1(2)
            TextBox4.Text = c1(3)
            TextBox3.Text = c1(4)
            TextBox2.Text = c1(5)
            TextBox1.Text = c1(6)
            'used to check wether there is an error and to display an error message
        Catch ex As FileNotFoundException
            Console.WriteLine("file does not exist")
        End Try



        Try
            'reading all the lines in the file 
            c2 = File.ReadAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c2.txt")
            'for loop to load the game  
            TextBox8.Text = c2(0)
            TextBox9.Text = c2(1)
            TextBox10.Text = c2(2)
            TextBox11.Text = c2(3)
            TextBox12.Text = c2(4)
            TextBox13.Text = c2(5)
            TextBox14.Text = c2(6)
            'Console.WriteLine(c2(i))
            'Next
            'used to check wether there is an error and to display an error message
        Catch ex As FileNotFoundException
            Console.WriteLine("file does not exist")
        End Try




        Try
            'reading all the lines in the file
            c3 = File.ReadAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c3.txt")
            'for loop to load the game 
            'For i = 0 To c3.Length - 1 Step 1
            TextBox15.Text = c3(0)
            TextBox16.Text = c3(1)
            TextBox17.Text = c3(2)
            TextBox18.Text = c3(3)
            TextBox19.Text = c3(4)
            TextBox20.Text = c3(5)
            TextBox21.Text = c3(6)
            'Next
            'used to check wether there is an error and to display an error message
        Catch ex As FileNotFoundException
            Console.WriteLine("file does not exist")
        End Try


        Try
            'reading all the lines in the file
            c4 = File.ReadAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c4.txt")
            'for loop to load the game 
            TextBox22.Text = c4(0)
            TextBox23.Text = c4(1)
            TextBox24.Text = c4(2)
            TextBox25.Text = c4(3)
            TextBox26.Text = c4(4)
            TextBox27.Text = c4(5)
            TextBox28.Text = c4(6)
            'used to check wether there is an error and to display an error message
        Catch ex As FileNotFoundException
            Console.WriteLine("file does not exist")
        End Try


        Try
            'reading all the lines in the file
            c5 = File.ReadAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c5.txt")
            'for loop to load the game

            TextBox29.Text = c5(0)
            TextBox30.Text = c5(1)
            TextBox31.Text = c5(2)
            TextBox32.Text = c5(3)
            TextBox33.Text = c5(4)
            TextBox34.Text = c5(5)
            TextBox35.Text = c5(6)
            'used to check wether there is an error and to display an error message
        Catch ex As FileNotFoundException
            Console.WriteLine("file does not exist")
        End Try


        Try
            'reading all the lines in the file
            c6 = File.ReadAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c6.txt")
            'for loop to load the game
            TextBox36.Text = c6(0)
            TextBox37.Text = c6(1)
            TextBox38.Text = c6(2)
            TextBox39.Text = c6(3)
            TextBox40.Text = c6(4)
            TextBox41.Text = c6(5)
            TextBox42.Text = c6(6)

            'used to check wether there is an error and to display an error message
        Catch ex As FileNotFoundException
            Console.WriteLine("file does not exist")
        End Try


        Try
            'reading all the lines in the file
            c7 = File.ReadAllLines("C:\Users\Chamaka\Desktop\visual basic\P00200362_Chamaka_Amarasinghe_ITP_Summer_Code\Assingment final\connectfour\c7.txt")
            'for loop to load the game
            TextBox43.Text = c7(0)
            TextBox44.Text = c7(1)
            TextBox45.Text = c7(2)
            TextBox46.Text = c7(3)
            TextBox47.Text = c7(4)
            TextBox48.Text = c7(5)
            TextBox49.Text = c7(6)

            'used to check wether there is an error and to display an error message
        Catch ex As FileNotFoundException
            Console.WriteLine("file does not exist")
        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        For a = 0 To 6
            'vbnullstring clears any previous values that have been stored
            c1(a) = vbNullString
        Next

        For a = 0 To 6
            'vbnullstring clears any previous values that have been stored
            c2(a) = vbNullString
        Next

        For a = 0 To 6
            'vbnullstring clears any previous values that have been stored
            c3(a) = vbNullString
        Next

        For a = 0 To 6
            'vbnullstring clears any previous values that have been stored
            c4(a) = vbNullString
        Next

        For a = 0 To 6
            'vbnullstring clears any previous values that have been stored
            c5(a) = vbNullString
        Next

        For a = 0 To 6
            'vbnullstring clears any previous values that have been stored
            c6(a) = vbNullString
        Next

        For a = 0 To 6
            'vbnullstring clears any previous values that have been stored
            c7(a) = vbNullString
        Next

        'assigning the null val to the text box
        TextBox7.Text = ""
        TextBox6.Text = ""
        TextBox5.Text = ""
        TextBox4.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        TextBox1.Text = ""

        'assigning the null val to the text box
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""

        'assigning the null val to the text box
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""

        'assigning the null val to the text box
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        TextBox28.Text = ""

        'assigning the null val to the text box
        TextBox29.Text = ""
        TextBox30.Text = ""
        TextBox31.Text = ""
        TextBox32.Text = ""
        TextBox33.Text = ""
        TextBox34.Text = ""
        TextBox35.Text = ""

        'assigning the null val to the text box
        TextBox36.Text = ""
        TextBox37.Text = ""
        TextBox38.Text = ""
        TextBox39.Text = ""
        TextBox40.Text = ""
        TextBox41.Text = ""
        TextBox42.Text = ""

        'assigning the null val to the text box
        TextBox43.Text = ""
        TextBox44.Text = ""
        TextBox45.Text = ""
        TextBox46.Text = ""
        TextBox47.Text = ""
        TextBox48.Text = ""
        TextBox49.Text = ""
    End Sub





End Class


