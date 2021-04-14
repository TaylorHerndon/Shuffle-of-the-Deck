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

        Dim shuffle As Boolean = False
        Dim firstCard As Boolean = True
        Dim repeat As Boolean = True
        Dim userShuffle As Boolean = False

        Dim suitName As String = ""
        Dim valueName As String = ""
        Dim deck(3, 12) As String
        Dim drawnLog(3, 12) As Boolean
        Dim collumn As Integer = 0 'Used to access the array containing card names
        Dim row As Integer = 0 'Used to access the array containing card names
        Dim collumn2 As Integer = 0 'Used to access the array containing draw history
        Dim row2 As Integer = 0 'Used to access the array containing draw history


        'Creating the deck of cards which is just a array of strings with the suit and value of each card.
        For collumn = 0 To 3

            For row = 0 To 12

                Select Case collumn
                    Case 0
                        suitName = "Hearts"
                    Case 1
                        suitName = "Diamonds"
                    Case 2
                        suitName = "Spades"
                    Case 3
                        suitName = "Clubs"
                End Select

                Select Case row
                    Case 0
                        valueName = "Ace"
                    Case 1
                        valueName = "2"
                    Case 2
                        valueName = "3"
                    Case 3
                        valueName = "4"
                    Case 4
                        valueName = "5"
                    Case 5
                        valueName = "6"
                    Case 6
                        valueName = "7"
                    Case 7
                        valueName = "8"
                    Case 8
                        valueName = "9"
                    Case 9
                        valueName = "10"
                    Case 10
                        valueName = "Jack"
                    Case 11
                        valueName = "Queen"
                    Case 12
                        valueName = "King"
                End Select

                deck(collumn, row) = valueName & " of " & suitName

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

        Do Until shuffle = True

            'Make sure the first card drawn is random
            If firstCard = True Then
                collumn = Convert.ToInt32(VBMath.Rnd * 3)
                row = Convert.ToInt32(VBMath.Rnd * 12)
            End If

            'Draw new card
            Console.Clear()
            Console.WriteLine("[Press any key to draw a new card]")

            If Console.ReadKey.Key = ConsoleKey.S Then
                userShuffle = True
            End If

            'Draw random cards until we reach a card that has not been drawn yet
            Do Until drawnLog(collumn, row) = False
                collumn = Convert.ToInt32(VBMath.Rnd * 3)
                row = Convert.ToInt32(VBMath.Rnd * 12)
            Loop

            'Mark the new card as drawn
            drawnLog(collumn, row) = True


            'Tell the user what card they drew
            If userShuffle = False Then
                Console.Clear()
                Console.WriteLine("You drew the " & deck(collumn, row) & ".")
                Console.WriteLine(vbNewLine & "[Press any key to continue]")

                If Console.ReadKey.Key = ConsoleKey.S Then
                    userShuffle = True
                End If
            End If

            'Check to see if all cards have been drawn
            collumn2 = 0
            row2 = 0

            Do Until drawnLog(collumn2, row2) = False Or repeat = False Or userShuffle <> False

                If row2 = 12 Then
                    row2 = 0
                    collumn2 = collumn2 + 1
                Else
                    row2 = row2 + 1
                End If

                If collumn2 = 3 And row2 = 12 Then
                    shuffle = True
                    repeat = False
                End If

            Loop

            repeat = True

            'If all cards have been drawn shuffle the deck or exit the program
            If shuffle = True Then

                Console.Clear()
                Thread.Sleep(1500)
                Console.WriteLine(vbNewLine & "That was the last card in the deck!")
                Thread.Sleep(1500)
                Console.WriteLine("Do you want to shuffle the deck and go again?")
                Thread.Sleep(500)
                Console.WriteLine("Y - N")

                Do Until repeat = False

                    Select Case Console.ReadKey.Key

                        Case ConsoleKey.Y
                            Console.WriteLine("Alright, here we go again!")
                            Thread.Sleep(2000)

                            For collumn2 = 0 To 3

                                For row2 = 0 To 12
                                    drawnLog(collumn2, row2) = False
                                Next

                            Next

                            repeat = False
                            shuffle = False

                        Case ConsoleKey.N
                            Console.Clear()
                            Console.WriteLine("Ok then, goodbye!")
                            Thread.Sleep(3000)
                            repeat = False
                            shuffle = True

                        Case Else
                            repeat = True

                    End Select

                Loop

                repeat = True

            End If

            'If the user presses [S] then ask the user if they are sure they want to reshuffle the deck.
            If userShuffle = True Then

                Console.Clear()
                Console.WriteLine("Oh! You said you want to reshuffle the deck are you sure?")
                Thread.Sleep(500)
                Console.WriteLine("Y - N")

                Do While repeat = True

                    Select Case Console.ReadKey.Key
                        Case ConsoleKey.Y
                            Console.Clear()
                            Console.WriteLine("Alrightly then!")
                            Thread.Sleep(1000)
                            Console.WriteLine("Anything you say boss!")
                            Thread.Sleep(1000)
                            Console.WriteLine("[Press any key to continue]")
                            Console.ReadKey()

                            For collumn2 = 0 To 3
                                For row2 = 0 To 12
                                    drawnLog(collumn2, row2) = False
                                Next
                            Next

                            userShuffle = False
                            repeat = False

                        Case ConsoleKey.N
                            Console.Clear()
                            Console.WriteLine("Ok, no problem!")
                            Thread.Sleep(1000)
                            Console.WriteLine("[Press any key to continue]")
                            Console.ReadKey()

                            userShuffle = False
                            repeat = False

                        Case Else
                            repeat = True

                    End Select

                Loop

                repeat = True

            End If

            firstCard = False

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
