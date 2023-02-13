'Joshua Makuch
'Spring 2023
'Impedance Calculator
'https://github.com/JoshuaMakuch/Impedance-Calculator.git

'Asks the user if it wants to do a parallel or series calculation, then it runs a function that then asks for how many impedances are in the network, once it has done that it
'asks the user to input the values in rectangular (real and agineryim) then stores those valuse into two arrays the size of the number of impedances

Option Explicit On
Option Strict On

Imports System
Imports System.Reflection.Metadata.Ecma335

Module ImpedanceCalcultor
    Sub Main()

        Dim userInput As String
        Dim totalImpedance() As String

        Do

            Console.WriteLine(vbCrLf & "Parallel or Series? S/P")
            userInput = Console.ReadLine()

            If userInput.ToLower() = "s" Then
                totalImpedance = Series()
                Console.Clear()
                For i = 0 To 3
                    Console.Write(totalImpedance(i) & " ")
                Next
            ElseIf userInput.ToLower() = "p" Then

            End If


        Loop

    End Sub

    Function Series() As String()

        Dim totalSeriesImpedance(4) As String
        Dim userInput As String

        Do
            Console.WriteLine("How many impedances?")
            userInput = Console.ReadLine
            Try
                Dim real(CInt(userInput)) As Double
                Dim fake(CInt(userInput)) As Double
                Console.WriteLine($"You chose {userInput} impedances.")
                Exit Do
            Catch ex As Exception
                Console.WriteLine("That is not a whole number")
            End Try
        Loop



        Return totalSeriesImpedance

    End Function

End Module



