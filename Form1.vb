Imports System.Windows.Forms.Design

Public Class Form1
    'Handle close button click
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub


    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        ' Validate that all fields have values before calculating
        If String.IsNullOrEmpty(cbLength.Text) OrElse
           String.IsNullOrEmpty(cbHeight.Text) OrElse
           String.IsNullOrEmpty(cbWidth.Text) OrElse
           String.IsNullOrEmpty(cbRollCoverage.Text) Then
            MessageBox.Show("Please enter all required values", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim length As Double
        Dim height As Double
        Dim width As Double
        Dim rollCoverage As Double

        'Check the data types of the values input (values can either be intergers or decimal)
        If Not Double.TryParse(cbLength.Text, length) OrElse length <= 0 Then
            MessageBox.Show("Please enter a valid positive number for Length", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not Double.TryParse(cbHeight.Text, height) OrElse height <= 0 Then
            MessageBox.Show("Please enter a valid positive number for Height", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not Double.TryParse(cbWidth.Text, width) OrElse width <= 0 Then
            MessageBox.Show("Please enter a valid positive number for Width", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not Double.TryParse(cbRollCoverage.Text, rollCoverage) OrElse rollCoverage <= 0 Then
            MessageBox.Show("Please enter a valid positive number for Roll Coverage", "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Calculate and display result
        Try
            Dim rolls As Integer = Calculator.CalculateNumberOfRollsPerRoom(length, width, height, rollCoverage)
            txtNumberOfSingleRolls.Text = rolls.ToString()
        Catch ex As Exception
            MessageBox.Show($"An error occurred during calculation: {ex.Message}", "Calculation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Event handlers to clear the result when any input changes
    Private Sub cbLength_TextChanged(sender As Object, e As EventArgs) Handles cbLength.TextChanged
        ClearResult()
    End Sub

    Private Sub cbHeight_TextChanged(sender As Object, e As EventArgs) Handles cbHeight.TextChanged
        ClearResult()
    End Sub

    Private Sub cbWidth_TextChanged(sender As Object, e As EventArgs) Handles cbWidth.TextChanged
        ClearResult()
    End Sub

    Private Sub cbRollCoverage_TextChanged(sender As Object, e As EventArgs) Handles cbRollCoverage.TextChanged
        ClearResult()
    End Sub

    'Utility Function to help clear the number of rolls on value changes
    Private Sub ClearResult()
        If Not String.IsNullOrWhiteSpace(txtNumberOfSingleRolls.Text) Then
            txtNumberOfSingleRolls.Clear()
        End If
    End Sub

End Class

' Separate the calculation logic in it's own function
Public Class Calculator
    Public Shared Function CalculateNumberOfRollsPerRoom(length As Double, width As Double, height As Double, rollCoverage As Double) As Integer
        ' Additional validation in case the method is called directly
        If length <= 0 OrElse width <= 0 OrElse height <= 0 OrElse rollCoverage <= 0 Then
            Throw New ArgumentException("All dimensions and roll coverage must be positive numbers")
        End If

        Dim totalArea As Double = 2 * (length * height) + 2 * (width * height)
        Dim numberOfRolls As Integer = CInt(Math.Ceiling(totalArea / rollCoverage))
        Return If(numberOfRolls < 1, 1, numberOfRolls) ' Ensure at least 1 roll is returned
    End Function
End Class
