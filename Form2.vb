Imports System.Drawing.Text

Public Class Form1

    'Handle close button click
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private depreciationBalances(,) As String = {}
    Public depreciationYearBalances As New List(Of Double())
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Setup depreciation balances columns
        depreciationBalances = {
            {1, 2000},
            {2, 1200},
            {3, 720},
            {4, 432},
            {5, 259}
        }

        lsbDecliningBalances.Items.Clear()
        lsbDecliningBalances.Items.Add("Year" & vbTab & "Depreciation")
        lsbSumOfYearDigits.Items.Add("Year" & vbTab & "Depreciation")
        For i As Integer = 0 To depreciationBalances.GetLength(0) - 1
            lsbDecliningBalances.Items.Add(depreciationBalances(i, 0) & vbTab & depreciationBalances(i, 1))
        Next

        If lsbDecliningBalances.Items.Count > 0 Then
            lsbDecliningBalances.SetSelected(1, True)
        End If


    End Sub

    Private Sub btnDisplayDepValues_Click(sender As Object, e As EventArgs) Handles btnDisplayDepValues.Click
        If String.IsNullOrEmpty(txtAssetCost.Text) OrElse
           String.IsNullOrEmpty(txtSalvageValue.Text) OrElse
           String.IsNullOrEmpty(cbUsefulLife.Text) Then
            MessageBox.Show("Please enter all values", "Missing Values", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim assetCost As Double
        Dim salvageValue As Double
        Dim usefulLife As Double

        If Not Double.TryParse(txtAssetCost.Text, assetCost) OrElse assetCost <= 0 Then
            MessageBox.Show("Please enter Asset Cost", "Invalid Input Value", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not Double.TryParse(txtSalvageValue.Text, salvageValue) OrElse salvageValue <= 0 Then
            MessageBox.Show("Please enter Salvage Value", "Invalid Input Value", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not Double.TryParse(cbUsefulLife.Text, usefulLife) OrElse usefulLife <= 0 Then
            MessageBox.Show("Please select Useful Life", "Invalid Input Value", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim selectedBalanceValue As String
        Dim selectedBalanceIndex As String = 1 ' set selected inde to 5 by default
        If (lsbDecliningBalances.SelectedIndex <> -1) Then
            selectedBalanceValue = lsbDecliningBalances.SelectedItem.ToString()
            selectedBalanceIndex = lsbDecliningBalances.SelectedIndex
        End If

        Try

            Dim value As Integer = Calculator.CalculateDepreciationValue(assetCost, salvageValue, usefulLife, selectedBalanceIndex)

            If value > 0 Then
                Dim newEntry As Double() = {selectedBalanceIndex, value}

                Dim isDuplicateEntry As Boolean = depreciationYearBalances.Any(Function(entry) entry(0) = selectedBalanceIndex)
                If Not isDuplicateEntry Then
                    depreciationYearBalances.Add(newEntry)
                    lsbSumOfYearDigits.Items.Add(newEntry(0) & vbTab & newEntry(1))
                Else
                    MessageBox.Show($"Year {selectedBalanceIndex} already calculated!", "Duplicate Entry")
                End If

            End If


        Catch ex As Exception
            MessageBox.Show($"An error occurred calculating: {ex.Message}", "Calculation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub cbUsefulLife_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbUsefulLife.SelectedIndexChanged
        lsbDecliningBalances.SelectedIndex = cbUsefulLife.Text
    End Sub
End Class


Public Class Calculator
    Public Shared Function CalculateDepreciationValue(assetCost As Double, salvageValue As Double, usefulLife As Double, period As Double)
        Dim depreciationValue As Double
        If assetCost <= 0 OrElse
           salvageValue <= 0 OrElse
           period <= 0 OrElse
           usefulLife <= 0 Then
            Throw New ArgumentException("Can not perfom calculation on 0 or non numbers")
        End If

        If usefulLife < period Then
            Throw New ArgumentException("Useflul life can canot be less than declining balance year(s)")
        Else
            depreciationValue = Financial.SYD(assetCost, salvageValue, usefulLife, period)
        End If

        Return FormatCurrency(Math.Round(depreciationValue, 2))

    End Function
End Class
