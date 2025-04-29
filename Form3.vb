Public Class Form1
    ' 2D array to store sales data: [region, month]
    Private salesData(,) As Decimal = New Decimal(2, 5) {
        {125000D, 190000D, 175000D, 188000D, 125000D, 163000D}, ' Kwazulu-Natal
        {90000D, 85000D, 80000D, 83000D, 87000D, 80000D},     ' Gauteng
        {65000D, 64000D, 71000D, 67000D, 65000D, 64000D}      ' Western Cape
    }

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set up the form
        Me.Text = "ITI Hub Regional Sales Report"
        Me.Size = New Size(800, 400)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Create controls
        CreateHeader()
        CreateSalesTable()
    End Sub

    Private Sub CreateHeader()
        Dim headerPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = Color.White
        }

        Dim lblTitle As New Label With {
            .Text = "Regional Sales Report (6 Months)",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point((Me.Width - 250) \ 2, 20)
        }
        headerPanel.Controls.Add(lblTitle)

        Me.Controls.Add(headerPanel)
    End Sub

    Private Sub CreateSalesTable()
        ' Create and configure DataGridView
        Dim dgvSales As New DataGridView With {
            .Location = New Point(20, 80),
            .Size = New Size(740, 250),
            .BorderStyle = BorderStyle.FixedSingle,
            .AllowUserToAddRows = False,
            .RowHeadersVisible = False,
            .ReadOnly = True
        }

        ' Style the header row
        dgvSales.EnableHeadersVisualStyles = False
        dgvSales.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
            .BackColor = Color.SteelBlue,
            .ForeColor = Color.White,
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .Alignment = DataGridViewContentAlignment.MiddleCenter
        }
        dgvSales.ColumnHeadersHeight = 30

        ' Add columns
        dgvSales.Columns.Add("Region", "Region")
        For month As Integer = 1 To 6
            dgvSales.Columns.Add($"Month{month}", $"Month {month}")
        Next
        dgvSales.Columns.Add("Percentage", "% Contribution")

        ' Format columns
        For Each col As DataGridViewColumn In dgvSales.Columns
            Select Case col.Name
                Case "Region"
                    col.Width = 120
                    col.DefaultCellStyle = New DataGridViewCellStyle With {
                        .Alignment = DataGridViewContentAlignment.MiddleLeft,
                        .Font = New Font("Arial", 9)
                    }
                Case "Percentage"
                    col.Width = 90
                    col.DefaultCellStyle = New DataGridViewCellStyle With {
                        .Format = "0\%",
                        .Alignment = DataGridViewContentAlignment.MiddleRight,
                        .Font = New Font("Arial", 9)
                    }
                Case Else
                    col.Width = 85
                    col.DefaultCellStyle = New DataGridViewCellStyle With {
                        .Format = "R#,##0",
                        .Alignment = DataGridViewContentAlignment.MiddleRight,
                        .Font = New Font("Arial", 9)
                    }
            End Select
        Next

        ' Calculate totals
        Dim companyTotal As Decimal = 0
        Dim regionTotals(2) As Decimal

        For region As Integer = 0 To 2
            For month As Integer = 0 To 5
                regionTotals(region) += salesData(region, month)
                companyTotal += salesData(region, month)
            Next
        Next

        ' Add data rows with accurate percentages
        Dim regions() As String = {"Kwazulu-Natal", "Gauteng", "Western Cape"}
        Dim percentages(2) As Integer
        Dim sumPct As Integer = 0

        For i As Integer = 0 To 1
            percentages(i) = CInt(Math.Round((regionTotals(i) / companyTotal) * 100))
            sumPct += percentages(i)
        Next
        percentages(2) = 100 - sumPct

        ' Add rows to DataGridView
        For i As Integer = 0 To 2
            Dim row = dgvSales.Rows.Add()
            dgvSales.Rows(row).Cells("Region").Value = regions(i)
            
            For month As Integer = 0 To 5
                dgvSales.Rows(row).Cells($"Month{month + 1}").Value = salesData(i, month)
            Next
            
            dgvSales.Rows(row).Cells("Percentage").Value = percentages(i)
        Next

        ' Add company total row
        Dim footerRow = dgvSales.Rows.Add()
        dgvSales.Rows(footerRow).Cells("Region").Value = "Company Total"
        
        For month As Integer = 0 To 5
            Dim monthTotal As Decimal = 0
            For region As Integer = 0 To 2
                monthTotal += salesData(region, month)
            Next
            dgvSales.Rows(footerRow).Cells($"Month{month + 1}").Value = monthTotal
        Next
        
        dgvSales.Rows(footerRow).Cells("Percentage").Value = 100
        dgvSales.Rows(footerRow).DefaultCellStyle = New DataGridViewCellStyle With {
            .Font = New Font(dgvSales.Font, FontStyle.Bold),
            .BackColor = Color.LightGray,
            .Alignment = DataGridViewContentAlignment.MiddleRight
        }
        dgvSales.Rows(footerRow).Cells("Region").Style.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.Controls.Add(dgvSales)
    End Sub
End Class
