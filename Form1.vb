Imports System.IO.Pipes
Imports System.Net.Mail
Imports System.Runtime.InteropServices
Imports System.Runtime.Remoting.Lifetime
Imports System.Runtime.Remoting.Messaging
Imports System.Security.Cryptography

Public Class GlobalVariables
    Public Shared carryOn As Boolean
End Class

Public Class Form1
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GlobalVariables.carryOn = True

        'clear the relevant textbox/hide label
        Label4.Hide()
        Label4.Text = ""
        TextBox2.Clear()

        'check for blank textbox ***** this wont work needs checking***
        If TextBox1.Text.Length < 1 Then
            Label4.Text = "Please enter a number!"
            GlobalVariables.carryOn = False
        End If

        'call a function to check for only numbers
        Dim pwLength As String
        Dim checkedOpt As String
        'set the variables to stop program bitching
        pwLength = TextBox1.Text
        checkedOpt = ""

        If GlobalVariables.carryOn = True Then
            'function to find out the length of password requested
            Call validNumLength(pwLength)
        End If
        If GlobalVariables.carryOn = True Then
            'find out what options have been chosen on the form
            Call checkedOptions(checkedOpt)
        End If

        If GlobalVariables.carryOn = True Then
            'create functino to loop for length of password at random
            Call PasswordCreater(checkedOpt, pwLength)
        End If



    End Sub

    Public Sub validNumLength(pwLength)
        GlobalVariables.carryOn = True
        Dim testLength As String
        testLength = ""

        'parse the text1, only numbers can be parsed
        If Integer.TryParse(pwLength, vbNull) Then
            'successful parse = number yay!
            testLength = pwLength
        Else
            'Parsing failed, let them know!
            Label4.Show()
            Label4.Text = "Please only enter numbers!"
            'Stop the program from continuing, gotta be a better way!
            GlobalVariables.carryOn = False
        End If
        'return the length required after the check
        pwLength = testLength

    End Sub

    Public Sub checkedOptions(ByRef checkedOpt As String)
        GlobalVariables.carryOn = True
        Dim randomString As String
        randomString = ""

        'check which options have been chosen, create a final string
        If CheckBox1.Checked Then
            randomString = randomString + "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        End If
        If CheckBox2.Checked Then
            randomString = randomString + "abcdefghijklmnopqrstuvwxyz"
        End If
        If CheckBox3.Checked Then
            randomString = randomString + "1234567890"
        End If
        If CheckBox4.Checked Then
            randomString = randomString + "!£$%^&*()_+=,.<>/?\"
        End If

        'check an option has been chosen, if not let them know!
        If CheckBox1.Checked = False And CheckBox2.Checked = False _
            And CheckBox3.Checked = False And CheckBox4.Checked = False Then
            Label4.Show()
            Label4.Text = "Please choose At Least one Option"
            GlobalVariables.carryOn = False
        End If

        'return the final string for useage
        checkedOpt = randomString
        Return

    End Sub

    Public Sub PasswordCreater(checkedOpt, pwLength)
        GlobalVariables.carryOn = True
        'create the password to the length of pwLength from the string of checkedOpt
        Dim rand As New Random
        Dim lettersCount As String
        Dim numLength As Integer
        Dim letters As String


        'save variable for easy reading
        letters = checkedOpt
        'need to know the length of option string
        lettersCount = checkedOpt.Length()
        'again easy reading attempted for loop
        numLength = pwLength

        'loop through time!
        Dim sb As New System.Text.StringBuilder
        For i As Integer = 1 To numLength
            Dim idx As Integer
            idx = rand.Next(0, lettersCount)
            sb.Append(letters.Substring(idx, 1))
        Next
        'place the result in the text box!
        TextBox2.Text = (sb.ToString())

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'simple exit the program jobby
        End
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'button to copy to clipboard with a notice added its completed
        Clipboard.SetText(TextBox2.Text)
        Label4.Show()
        Label4.Text = "Copied to clipboard!"
    End Sub
End Class
