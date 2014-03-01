RockPaperScissors
=================
Module HangMan

    Dim LettersGuessedArray(25, 1) As String 'AN ARRAY THAT WILL HOLD THE VALUES OF THE LETTERS GUESSED
    Dim LetterToGuess As String
    Dim LetterGuessLen As Integer
    Dim WordToBeGuessed As String
    Dim WordToBeGuessedCenc As String = "*"
    Dim WordLetter As String
    Dim checked As Boolean
    Dim lenX As Integer = 0 'DEFAULT VALUE = 0
    Dim letterVal As Integer
    Dim turnLimit As Integer
    Dim letterStr As String
    Dim PlayAgain As Integer
    Dim StrPlayAgain As String
    Sub Main()

        SetupConsole()
        Console.WriteLine("Welcome to HangMan")
        Console.WriteLine("Programmed by Keiron Segger")
        Console.WriteLine("Contact me at reddit.com/u/Segger96")
        Console.WriteLine("Have fun please report any bugs on the thread and feedback")
        lenX = Len(WordToBeGuessed)
        Do
            WordToBeGuessedCenc = WordToBeGuessedCenc + "*"
        Loop Until Len(WordToBeGuessedCenc) = lenX

        Console.WriteLine("The Word is " & WordToBeGuessedCenc)
        Do
            Do
                Console.WriteLine("Please enter a single letter to be guessed")
                LetterToGuess = Console.ReadLine()
                LetterToGuess = LCase(LetterToGuess)
                checked = ValidLetter(LetterToGuess)
                LetterGuessLen = Len(LetterToGuess)
                If checked = False Then

                ElseIf checked = True And LetterGuessLen = 1 Then
                    Exit Do
                End If
            Loop Until checked = True
            For x = 0 To 25
                If LettersGuessedArray(x, 0) = LetterToGuess Then
                    letterVal = x
                    'LettersGuessedArray(letterVal, 1) = 1
                End If
            Next
            For wordFind = 1 To lenX
                WordLetter = Mid(WordToBeGuessed, wordFind, 1)
                WordLetter = LCase(WordLetter)
                If LetterToGuess = WordLetter Then
                    Mid(WordToBeGuessedCenc, wordFind, 1) = LetterToGuess
                End If
            Next
            Console.WriteLine(WordToBeGuessedCenc)
            If LettersGuessedArray(letterVal, 1) = 0 Then
                turnLimit += 1
                LettersGuessedArray(letterVal, 1) = 1
            End If
            Console.WriteLine("Turns Taken = {0}/10", turnLimit)
            WordToBeGuessed = LCase(WordToBeGuessed)
            WordToBeGuessedCenc = LCase(WordToBeGuessedCenc)
        Loop Until WordToBeGuessedCenc = WordToBeGuessed Or turnLimit = 10
        If turnLimit = 10 Then
            MsgBox("You Lose")
        Else
            MsgBox("Congratulations You Won")
        End If
        Console.WriteLine("Do you want to play again??")
        StrPlayAgain = Console.ReadLine()



        If StrPlayAgain = "yes" Or StrPlayAgain = "y" Then
            Main()
        ElseIf StrPlayAgain = "no" Or StrPlayAgain = "n" Then
            Exit Sub
        Else
            Console.WriteLine("Do you want to play again??")
            StrPlayAgain = Console.ReadLine()
            StrPlayAgain = LCase(StrPlayAgain)
        End If


    End Sub
    Sub SetupConsole()
        Dim Checker As Integer = 0
        Dim CharToChck As String
        Dim WordLen As Integer
        Dim CheckerCount As Integer = 0
        Do
            Checker = 0 'WITHOUT THIS IT WILL CAUSE A CONSTANT LOOP AS IT DOES NOT RE-READ THE TOP VALUES
            CheckerCount = 0 'ALLOWS THE PROGRAM TO NOT LOOP THE SECOND TIME SETUPCONSOLE IS RUN, EG NUMBER ENTERED
            Console.WriteLine("Please enter the word to be guessed")
            Console.WriteLine("This is not AlphaNumeric, letters only please")
            WordToBeGuessed = Console.ReadLine()


            WordLen = Len(WordToBeGuessed)
            For X = 1 To WordLen
                CharToChck = Mid(WordToBeGuessed, X, 1)

                If IsNumeric(CharToChck) = True Then
                    Checker = 1
                Else
                    CheckerCount += 1
                End If
            Next
        Loop Until CheckerCount = WordLen And Checker < 1
        For LetterCount = 0 To 25
            For x = 0 To 1
                LettersGuessedArray(LetterCount, 0) = Chr(97 + LetterCount)
                Console.WriteLine(LettersGuessedArray(LetterCount, 0))
                LettersGuessedArray(LetterCount, 0) = LCase(LettersGuessedArray(LetterCount, 0))
            Next
        Next

        Console.Clear()

    End Sub
    Private Function ValidLetter(ByVal Letter As String) As Boolean
        If IsNumeric(Letter) = False Then
            Return True
        Else
            Return False
        End If
    End Function
End Module
