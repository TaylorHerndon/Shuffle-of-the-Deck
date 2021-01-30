Option Strict On
Option Explicit On

Imports System.Threading

'Taylor Herndon
'RCET0265
'Spring 2021
'Shuffle of the Deck
'https://github.com/TaylorHerndon/Shuffle-of-the-Deck
Module ShuffleOfTheDeck

    Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Integer) As Boolean

    Sub Main()

        Dim thread1 = New Thread(AddressOf DeckProgram)
        Dim thread2 = New Thread(AddressOf Abort)

        thread1.Start()
        thread2.Start()

    End Sub

    Sub DeckProgram()

        Randomize()

        Dim Shuffle As Boolean = False
        Dim FirstCard As Boolean = True
        Dim Repeat As Boolean = True
        Dim UserShuffle As Boolean = False

        Dim SuitName As String = ""
        Dim ValueName As String = ""
        Dim Deck(3, 12) As String
        Dim DrawnLog(3, 12) As Boolean
        Dim c As Integer = 0
        Dim r As Integer = 0
        Dim C2 As Integer = 0
        Dim R2 As Integer = 0


        'Creating the deck of cards which is just a array of strings with the suit and value of each card.
        For c = 0 To 3

            For r = 0 To 12

                Select Case c

                    Case 0
                        SuitName = "Hearts"
                    Case 1
                        SuitName = "Diamonds"
                    Case 2
                        SuitName = "Spades"
                    Case 3
                        SuitName = "Clubs"

                End Select

                Select Case r

                    Case 0
                        ValueName = "Ace"
                    Case 1
                        ValueName = "2"
                    Case 2
                        ValueName = "3"
                    Case 3
                        ValueName = "4"
                    Case 4
                        ValueName = "5"
                    Case 5
                        ValueName = "6"
                    Case 6
                        ValueName = "7"
                    Case 7
                        ValueName = "8"
                    Case 8
                        ValueName = "9"
                    Case 9
                        ValueName = "10"
                    Case 10
                        ValueName = "Jack"
                    Case 11
                        ValueName = "Queen"
                    Case 12
                        ValueName = "King"

                End Select

                Deck(c, r) = ValueName & " of " & SuitName

            Next

        Next

        'Introduction sequence
        Console.WriteLine("Hello!")
        Thread.Sleep(1000)
        Console.WriteLine("Welcome to Shuffle-of-the-Deck(TM)!")
        Thread.Sleep(2000)
        Console.WriteLine("Here you will be drawing from a shuffled deck of cards!")
        Thread.Sleep(2000)
        Console.WriteLine("At any point you can press [S] to reshuffle your deck of cards!")
        Thread.Sleep(2000)
        Console.WriteLine("You can also press [Q] to exit the program at any time!")
        Thread.Sleep(2000)
        Console.WriteLine("Without further a do lets start drawing cards!")
        Thread.Sleep(1500)
        Console.WriteLine("[Press any key to continue]")
        Console.ReadKey()
        Console.Clear()

        Do Until Shuffle = True

            'Make sure the first card drawn is random
            If FirstCard = True Then
                c = Convert.ToInt32(VBMath.Rnd * 3)
                r = Convert.ToInt32(VBMath.Rnd * 12)
            End If

            'Draw new card
            Console.Clear()
            Console.WriteLine("[Press any key to draw a new card]")

            If Console.ReadKey.Key = ConsoleKey.S Then
                UserShuffle = True
            End If

            'Draw random cards until we reach a card that has not been drawn yet
            Do Until DrawnLog(c, r) = False

                c = Convert.ToInt32(VBMath.Rnd * 3)
                r = Convert.ToInt32(VBMath.Rnd * 12)

            Loop

            'Mark the new card as drawn
            DrawnLog(c, r) = True


            'Tell the user what card they drew
            If UserShuffle = False Then

                Console.Clear()
                Console.WriteLine("You drew the " & Deck(c, r) & ".")
                Console.WriteLine(vbNewLine & "[Press any key to continue]")

                If Console.ReadKey.Key = ConsoleKey.S Then
                    UserShuffle = True
                End If

            End If


            'Check to see if all cards have been drawn
            C2 = 0
            R2 = 0

            Do Until DrawnLog(C2, R2) = False Or Repeat = False Or UserShuffle <> False

                If R2 = 12 Then

                    R2 = 0
                    C2 = C2 + 1

                Else

                    R2 = R2 + 1

                End If

                If C2 = 3 And R2 = 12 Then

                    Shuffle = True
                    Repeat = False

                End If

            Loop

            Repeat = True

            'If all cards have been drawn shuffle the deck or exit the program
            If Shuffle = True Then

                Console.Clear()
                Thread.Sleep(1500)
                Console.WriteLine(vbNewLine & "That was the last card in the deck!")
                Thread.Sleep(1500)
                Console.WriteLine("Do you want to shuffle the deck and go again?")
                Thread.Sleep(500)
                Console.WriteLine("Y - N")

                Do Until Repeat = False

                    Select Case Console.ReadKey.Key

                        Case ConsoleKey.Y

                            Console.WriteLine("Alright, here we go again!")
                            Thread.Sleep(2000)

                            For C2 = 0 To 3

                                For R2 = 0 To 12

                                    DrawnLog(C2, R2) = False

                                Next

                            Next

                            Repeat = False
                            Shuffle = False

                        Case ConsoleKey.N

                            Console.Clear()
                            Console.WriteLine("Ok then, goodbye!")
                            Thread.Sleep(3000)
                            Repeat = False
                            Shuffle = True

                        Case Else

                            Repeat = True

                    End Select

                Loop

                Repeat = True

            End If

            'If the user presses [S] then ask the user if they are sure they want to reshuffle the deck.
            If UserShuffle = True Then

                Console.Clear()
                Console.WriteLine("Oh! You said you want to reshuffle the deck are you sure?")
                Thread.Sleep(500)
                Console.WriteLine("Y - N")

                Do While Repeat = True

                    Select Case Console.ReadKey.Key
                        Case ConsoleKey.Y

                            Console.Clear()
                            Console.WriteLine("Alrightly then!")
                            Thread.Sleep(1000)
                            Console.WriteLine("Anything you say boss!")
                            Thread.Sleep(1000)
                            Console.WriteLine("[Press any key to continue]")
                            Console.ReadKey()

                            For C2 = 0 To 3

                                For R2 = 0 To 12

                                    DrawnLog(C2, R2) = False

                                Next

                            Next

                            UserShuffle = False
                            Repeat = False

                        Case ConsoleKey.N

                            Console.Clear()
                            Console.WriteLine("Ok, no problem!")
                            Thread.Sleep(1000)
                            Console.WriteLine("[Press any key to continue]")
                            Console.ReadKey()

                            UserShuffle = False
                            Repeat = False

                        Case Else

                            Repeat = True

                    End Select

                Loop

                Repeat = True

            End If

            FirstCard = False

        Loop

    End Sub

    Sub Abort()

        'Seprate thread to exit the program 
        'Added this because the program takes a really long time before it asks you if you want to exit

        Do Until GetAsyncKeyState(81) = True
            Thread.Sleep(100)
        Loop

        End

    End Sub

End Module
