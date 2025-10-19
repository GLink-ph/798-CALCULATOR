Imports System.Resources.ResXFileRef

Public Class Form1
    ' Conversion table for fraction values
    Private FractionValues As Dictionary(Of String, Double) = New Dictionary(Of String, Double) From {
        {"0", 0}, {"1/16", 1 / 16}, {"1/8", 1 / 8}, {"3/16", 3 / 16},
        {"1/4", 1 / 4}, {"5/16", 5 / 16}, {"3/8", 3 / 8}, {"7/16", 7 / 16},
        {"1/2", 1 / 2}, {"9/16", 9 / 16}, {"5/8", 5 / 8}, {"11/16", 11 / 16},
        {"3/4", 3 / 4}, {"13/16", 13 / 16}, {"7/8", 7 / 8}, {"15/16", 15 / 16}
    }

    Private Function ConvertToDouble(value As String) As Double
        ' Trim any leading or trailing whitespace
        Dim trimmedValue As String = value.Trim()

        ' Split the value by space to separate whole number and fraction
        Dim parts() As String = trimmedValue.Split(" "c)

        ' Initialize variables for the whole number and fraction parts
        Dim whole As Double = 0
        Dim fraction As Double = 0

        ' If there are two parts (whole number and fraction)
        If parts.Length = 2 Then
            ' Attempt to parse the whole number part
            If Double.TryParse(parts(0), whole) Then
                ' Split the fraction part by '/' to get numerator and denominator
                Dim fractionParts() As String = parts(1).Split("/"c)

                ' Check if both parts of the fraction are present
                If fractionParts.Length = 2 Then
                    ' Attempt to parse the numerator and denominator
                    Dim numerator As Double
                    Dim denominator As Double

                    If Double.TryParse(fractionParts(0), numerator) AndAlso Double.TryParse(fractionParts(1), denominator) Then
                        ' Calculate the fraction value
                        fraction = numerator / denominator
                    End If
                End If
            End If
        ElseIf parts.Length = 1 Then ' If there's only one part
            ' Attempt to parse it as a whole number or fraction
            Dim fractionParts() As String = parts(0).Split("/"c)

            ' Check if it's a mixed fraction
            If fractionParts.Length = 2 Then
                Dim wholePart As Double
                Dim numerator As Double
                Dim denominator As Double

                If Double.TryParse(fractionParts(0), wholePart) AndAlso Double.TryParse(fractionParts(1), numerator) AndAlso Double.TryParse(parts(0), denominator) Then
                    ' Convert mixed fraction to decimal
                    fraction = (wholePart + (numerator / denominator))
                End If
            Else
                ' If it's not a mixed fraction, attempt to parse it as a decimal
                If Double.TryParse(parts(0), whole) Then
                    fraction = 0
                End If
            End If
        End If

        ' Return the sum of whole and fraction
        Return whole + fraction
    End Function

    ' Function to convert decimal to whole number and fraction
    Private Function ConvertToFraction(value As Double) As String
        Dim wholeNumber As Integer = Math.Floor(value)
        Dim fractionalPart As Double = value - wholeNumber

        ' Check if the exact fractional value exists in the dictionary
        If FractionValues.ContainsValue(fractionalPart) Then
            ' If found, retrieve its corresponding key (fraction)
            Dim fraction As String = FractionValues.FirstOrDefault(Function(kv) kv.Value = fractionalPart).Key

            ' Combine whole number and fraction
            If wholeNumber = 0 AndAlso fraction = "0" Then
                Return "0"
            ElseIf fraction = "0" Then
                Return wholeNumber.ToString()
            ElseIf wholeNumber = 0 Then
                Return fraction
            Else
                Return wholeNumber.ToString() & " " & fraction
            End If
        Else
            ' Round the fractional part to the nearest fraction with denominator 16
            Dim roundedFraction As Double = Math.Round(fractionalPart * 16) / 16

            ' Check if the rounded fraction is 1 (whole)
            If roundedFraction = 1 Then
                wholeNumber += 1
                roundedFraction = 0
            End If

            ' Combine whole number and rounded fraction
            If wholeNumber = 0 AndAlso roundedFraction = 0 Then
                Return "0"
            ElseIf roundedFraction = 0 Then
                Return wholeNumber.ToString()
            ElseIf wholeNumber = 0 Then
                Return FractionValues.FirstOrDefault(Function(kv) kv.Value = roundedFraction).Key
            Else
                Return wholeNumber.ToString() & " " & FractionValues.FirstOrDefault(Function(kv) kv.Value = roundedFraction).Key
            End If
        End If
    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ' Retrieve height and length inputs and convert to double
            Dim height As Double = ConvertToDouble(TextBox1.Text)
            Dim length As Double = ConvertToDouble(TextBox2.Text)

            ' Perform calculations
            Dim doubleHead As Double = length - 1.125 ' 1 1/8 inches is 1.125 in decimal
            Dim doubleSill As Double = length - 1.125
            Dim doubleJamb As Double = height
            Dim interLocker As Double = height - 1.75 ' 1 3/4 inches is 1.75 in decimal
            Dim lockStyle As Double = height - 1.75
            Dim topBottomRail As Double = (length - 5) / 2

            ' Convert results to whole number and fraction
            TextBox3.Text = ConvertToFraction(doubleHead)
            TextBox4.Text = ConvertToFraction(doubleSill)
            TextBox5.Text = ConvertToFraction(doubleJamb)
            TextBox8.Text = ConvertToFraction(interLocker)
            TextBox7.Text = ConvertToFraction(lockStyle)
            TextBox6.Text = ConvertToFraction(topBottomRail)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub


End Class
