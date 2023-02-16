'Joshua Makuch
'Spring 2023
'Impedance Calculator
'https://github.com/JoshuaMakuch/Impedance-Calculator.git

'Asks the user if it wants to do a parallel or series calculation, then it runs a function that then asks for how many impedances are in the network, once it has done that it
'asks the user to input the values in rectangular (real and fake) then stores those valuse into two arrays the size of the number of impedances

Option Explicit On
Option Strict On

Imports System
Imports System.Math

Module ImpedanceCalcultor
    Sub Main()

        Dim userInput As String
        Dim totalImpedance() As Double

        Do
            Console.WriteLine("Parallel or Series? S/P")
            userInput = Console.ReadLine()

            If userInput.ToLower() = "s" Then

                totalImpedance = Series()
                Console.Clear()
                Console.WriteLine($"Real: {totalImpedance(0)} | Imaginary: {totalImpedance(1)}")
                Console.WriteLine(vbCrLf & $"Polar: {totalImpedance(2)} | Angle: {totalImpedance(3)}°" & vbCrLf)

            ElseIf userInput.ToLower() = "p" Then

                totalImpedance = Parallel()
                'Console.Clear()
                Console.WriteLine($"Real: {totalImpedance(0)} | Imaginary: {totalImpedance(1)}")
                Console.WriteLine(vbCrLf & $"Polar: {totalImpedance(2)} | Angle: {totalImpedance(3)}°" & vbCrLf)

            End If

        Loop

    End Sub



    Function Series() As Double()

        Dim totalSeriesImpedance(4) As Double
        Dim impedances(0, 0) As Double
        Dim userInput As String

        'ReDims an array containing however many impedances the user chooses
        Do
            Console.WriteLine(vbCrLf & "How many impedances?: ")
            userInput = Console.ReadLine
            Try
                ReDim impedances(Math.Abs(CInt(userInput)) - 1, 1)
                Console.WriteLine(vbCrLf & $"You chose {Math.Abs(CInt(userInput))} impedances.")
                Exit Do
            Catch ex As Exception
                Console.WriteLine("That is not a whole number")
            End Try
        Loop

        'Populates the array's values based off of user's inputs
        For i = 0 To impedances.GetLength(0) - 1
            For j = 0 To 1
                Select Case j
                    Case 0
                        Console.WriteLine(vbCrLf & $"What is the real of impedance #{i + 1}")
                        impedances(i, j) = CDbl(Console.ReadLine())
                    Case 1
                        Console.WriteLine($"What is the imaginary of impedance #{i + 1}")
                        impedances(i, j) = CDbl(Console.ReadLine())
                End Select
            Next
        Next

        'Adds the series impedances into an array to return back to the main code to be displayed
        For i = 0 To impedances.GetLength(0) - 1
            For j = 0 To 1
                Select Case j
                    Case 0
                        totalSeriesImpedance(0) = totalSeriesImpedance(0) + impedances(i, j)
                    Case 1
                        totalSeriesImpedance(1) = totalSeriesImpedance(1) + impedances(i, j)
                End Select
            Next
        Next

        'Converts the two rectangular impedances into polar and stores them to be sent out
        totalSeriesImpedance(2) = ((totalSeriesImpedance(0)) ^ 2 + totalSeriesImpedance(1) ^ 2) ^ 0.5
        totalSeriesImpedance(3) = (Atan(totalSeriesImpedance(1) / totalSeriesImpedance(0)) * 360) / (2 * Math.PI)

        Return totalSeriesImpedance

    End Function

    Function Parallel() As Double()

        Dim totalParallelImpedance(4) As Double
        Dim tempStorageOfImpedances(4) As Double
        Dim impedances(0, 0) As Double
        Dim userInput As String

        'ReDims an array containing however many impedances the user chooses
        Do
            Console.WriteLine(vbCrLf & "How many impedances?: ")
            userInput = Console.ReadLine
            Try
                ReDim impedances(Math.Abs(CInt(userInput)) - 1, 1)
                Console.WriteLine(vbCrLf & $"You chose {Math.Abs(CInt(userInput))} impedances.")
                Exit Do
            Catch ex As Exception
                Console.WriteLine("That is not a whole number")
            End Try
        Loop

        'Populates the array's values based off of user's inputs
        For i = 0 To impedances.GetLength(0) - 1
            For j = 0 To 1
                Select Case j
                    Case 0
                        Console.WriteLine(vbCrLf & $"What is the real of impedance #{i + 1}")
                        impedances(i, j) = CDbl(Console.ReadLine())
                    Case 1
                        Console.WriteLine($"What is the imaginary of impedance #{i + 1}")
                        impedances(i, j) = CDbl(Console.ReadLine())
                End Select
            Next
        Next

        'Populates he first portion of the array to be an open ensure that the first stored value isnt a short
        totalParallelImpedance(2) = 1 * 10 ^ 9
        totalParallelImpedance(3) = 0

        'Major Problem, VB is in radians, not degrees...

        'Adds the parallel impedances into an array to return back to the main code to be displayed
        For i = 0 To impedances.GetLength(0) - 1
            'Stores the inverse of so far total polar impedance to be parallel'd
            tempStorageOfImpedances(0) = 1 / totalParallelImpedance(2)
            tempStorageOfImpedances(1) = -totalParallelImpedance(3)
            Console.WriteLine(tempStorageOfImpedances(0))
            Console.WriteLine(tempStorageOfImpedances(1))
            'Stores the inverse of the next impedance ready to be parallel'd 
            tempStorageOfImpedances(2) = 1 / ((impedances(i, 0) ^ 2 + impedances(i, 1) ^ 2) ^ 0.5)
            tempStorageOfImpedances(3) = -(Atan(impedances(i, 1) / impedances(i, 0)))
            Console.WriteLine(tempStorageOfImpedances(2))
            Console.WriteLine(tempStorageOfImpedances(3))
            'Combines the two ready to be parallel'd values into polar
            totalParallelImpedance(2) = 1 / ((((tempStorageOfImpedances(0) * Cos(tempStorageOfImpedances(1))) +
                (tempStorageOfImpedances(2) * Cos(tempStorageOfImpedances(3)))) ^ 2 +
                ((tempStorageOfImpedances(0) * Sin(tempStorageOfImpedances(1))) +
                (tempStorageOfImpedances(2) * Sin(tempStorageOfImpedances(3)))) ^ 2) ^ 0.5)
            'This is for the phase angle of the two combined values
            totalParallelImpedance(3) = -Atan(((tempStorageOfImpedances(0) * Sin(tempStorageOfImpedances(1))) +
                (tempStorageOfImpedances(2) * Sin(tempStorageOfImpedances(3)))) /
                ((tempStorageOfImpedances(0) * Cos(tempStorageOfImpedances(1))) +
                (tempStorageOfImpedances(2) * Cos(tempStorageOfImpedances(3)))))
            Console.WriteLine(totalParallelImpedance(2))
            Console.WriteLine(totalParallelImpedance(3))
        Next

        'Converts the two polar impedances values into rectangular and stores them to be sent out
        totalParallelImpedance(0) = totalParallelImpedance(2) * Cos(totalParallelImpedance(3))
        totalParallelImpedance(1) = totalParallelImpedance(2) * Sin(totalParallelImpedance(3))
        totalParallelImpedance(3) = totalParallelImpedance(3) * 360 / (2 * Math.PI)

        Return totalParallelImpedance

    End Function

End Module



