' Project:          Piecework V2
' Programmer:       Matthew Lane
' Description:      This project calculates the pay based on a scale for pieceworkers includes error catching.



Public Class Form1
    'Declare variables and constants.

    Private AmountOfOrders As Integer
    Private AmountOfPiecesDecimal As Decimal
    Private Total, AmountDue, AverageDecimal As Decimal
    Const ONETONINETYNINEDecimal As Decimal = 0.5
    Const TWOHUNDREDTOTHREEDecimal As Decimal = 0.55
    Const FOURHUNDREDTOFIVEDecimal As Decimal = 0.6
    Const SIXHUNDREDORMOREDecimal As Decimal = 0.65

    Private Function CalculatePay() As Decimal
        ' Calculate the Amount Due.

        Select Case AmountOfPiecesDecimal
            Case 1 To 199
                AmountDue = AmountOfPiecesDecimal * ONETONINETYNINEDecimal
            Case 200 To 399
                AmountDue = AmountOfPiecesDecimal * TWOHUNDREDTOTHREEDecimal
            Case 400 To 599
                AmountDue = AmountOfPiecesDecimal * FOURHUNDREDTOFIVEDecimal
            Case Is >= 600
                AmountDue = AmountOfPiecesDecimal * SIXHUNDREDORMOREDecimal
        End Select
        Return AmountDue

    End Function

    Private Sub CalculatePayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculatePayToolStripMenuItem.Click
        'Calculate and display the totals and add to summary

        'Error catching for Amount of Peices.
        Try
            AmountOfPiecesDecimal = Decimal.Parse(PiecesTextBox.Text)
        Catch ex As Exception
            MessageBox.Show("Amount must be numeric", "Data entry error", MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        End Try

        'Format and display the amount due in the appropriate text box.
        AmountEarnedLabel.Text = CalculatePay.ToString("C")

        'Add to totals
        AmountOfOrders += 1
        Total += AmountDue

        SummaryToolStripMenuItem.Enabled = True
    End Sub

    Private Sub SummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SummaryToolStripMenuItem.Click
        ' Display summary information.

        Dim MessageString As String

        If Total > 0 Then
            'Calculate the average
            AverageDecimal = Total / AmountOfOrders

            'Concatenate the message string
            MessageString = "Number of payouts: " &
                AmountOfOrders.ToString() & Environment.NewLine &
                "Average amount paid: " & AverageDecimal.ToString("C") &
                Environment.NewLine
            MessageBox.Show(MessageString, "Payout Summary", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Else
            MessageString = "No data to summarize."
            MessageBox.Show(MessageString, "Number of Payouts", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        'Closes the program.

        Me.Close()
    End Sub

    Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        ' Clear the text boxes but not delete summary data.

        AmountEarnedLabel.Text = ""
        PiecesTextBox.Clear()
        With NameTextBox
            .Clear()
            .Focus()
        End With
    End Sub

    Private Sub ClearAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllToolStripMenuItem.Click
        ' Clear everything including summary information.

        Dim MessageString As String
        Dim ReturnDialogResult As DialogResult

        MessageString = "Are you sure?" & Environment.NewLine & "Clear all Deletes Summary Data"
        ReturnDialogResult = MessageBox.Show(MessageString, "Clear All", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk)


        If ReturnDialogResult = Windows.Forms.DialogResult.Yes Then
            AverageDecimal = 0
            AmountOfOrders = 0
            Total = 0
            AmountEarnedLabel.Text = ""
            PiecesTextBox.Clear()
            With NameTextBox
                .Clear()
                .Focus()
            End With
            SummaryToolStripMenuItem.Enabled = False
        End If

    End Sub

    Private Sub ColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ColorToolStripMenuItem.Click
        ' Allow the user to select a new color for the amount earned.

        With ColorDialog1
            .Color = AmountEarnedLabel.ForeColor
            .ShowDialog()
            AmountEarnedLabel.ForeColor = .Color
        End With
    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontToolStripMenuItem.Click
        'Change the font of a number?

        With FontDialog1
            .Font = AmountEarnedLabel.Font
            .ShowDialog()
            AmountEarnedLabel.Font = .Font
        End With
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        'Display the about message box.

        Dim MessageString As String

        MessageString = "Piecework Pay" & Environment.NewLine & "Programmed by Matthew Lane"
        MessageBox.Show(MessageString, "About Piecework Pay", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class