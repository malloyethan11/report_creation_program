

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnCreditDepositReport = New System.Windows.Forms.Button()
        Me.btnCashDepositReport = New System.Windows.Forms.Button()
        Me.btnInventoryReport = New System.Windows.Forms.Button()
        Me.btnTaxReport = New System.Windows.Forms.Button()
        Me.btnSalesReport = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnCreditDepositReport
        '
        Me.btnCreditDepositReport.Location = New System.Drawing.Point(54, 290)
        Me.btnCreditDepositReport.Name = "btnCreditDepositReport"
        Me.btnCreditDepositReport.Size = New System.Drawing.Size(135, 46)
        Me.btnCreditDepositReport.TabIndex = 0
        Me.btnCreditDepositReport.Text = "Credit Deposit Report"
        Me.btnCreditDepositReport.UseVisualStyleBackColor = True
        '
        'btnCashDepositReport
        '
        Me.btnCashDepositReport.Location = New System.Drawing.Point(54, 225)
        Me.btnCashDepositReport.Name = "btnCashDepositReport"
        Me.btnCashDepositReport.Size = New System.Drawing.Size(135, 46)
        Me.btnCashDepositReport.TabIndex = 1
        Me.btnCashDepositReport.Text = "Cash Deposit Report"
        Me.btnCashDepositReport.UseVisualStyleBackColor = True
        '
        'btnInventoryReport
        '
        Me.btnInventoryReport.Location = New System.Drawing.Point(54, 157)
        Me.btnInventoryReport.Name = "btnInventoryReport"
        Me.btnInventoryReport.Size = New System.Drawing.Size(135, 46)
        Me.btnInventoryReport.TabIndex = 2
        Me.btnInventoryReport.Text = "Inventory Report"
        Me.btnInventoryReport.UseVisualStyleBackColor = True
        '
        'btnTaxReport
        '
        Me.btnTaxReport.Location = New System.Drawing.Point(54, 93)
        Me.btnTaxReport.Name = "btnTaxReport"
        Me.btnTaxReport.Size = New System.Drawing.Size(135, 46)
        Me.btnTaxReport.TabIndex = 3
        Me.btnTaxReport.Text = "Tax Report"
        Me.btnTaxReport.UseVisualStyleBackColor = True
        '
        'btnSalesReport
        '
        Me.btnSalesReport.Location = New System.Drawing.Point(54, 29)
        Me.btnSalesReport.Name = "btnSalesReport"
        Me.btnSalesReport.Size = New System.Drawing.Size(135, 46)
        Me.btnSalesReport.TabIndex = 4
        Me.btnSalesReport.Text = "Sales Report"
        Me.btnSalesReport.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(54, 376)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(135, 46)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(51, 345)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(139, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "______________________"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(242, 434)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSalesReport)
        Me.Controls.Add(Me.btnTaxReport)
        Me.Controls.Add(Me.btnInventoryReport)
        Me.Controls.Add(Me.btnCashDepositReport)
        Me.Controls.Add(Me.btnCreditDepositReport)
        Me.Name = "frmMain"
        Me.Text = "Home"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnCreditDepositReport As Button
    Friend WithEvents btnCashDepositReport As Button
    Friend WithEvents btnInventoryReport As Button
    Friend WithEvents btnTaxReport As Button
    Friend WithEvents btnSalesReport As Button




    Private Sub btnSalesReport_Click(sender As Object, e As EventArgs) Handles btnSalesReport.Click

        ' declare variables
        Dim frmSalesReportInfo = New frmSalesReportInfo()

        ' open intermediate form to get email address and time period
        frmSalesReportInfo.ShowDialog()

    End Sub

    Private Sub btnTaxReport_Click(sender As Object, e As EventArgs) Handles btnTaxReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

    End Sub

    Private Sub btnInventoryReport_Click(sender As Object, e As EventArgs) Handles btnInventoryReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

    End Sub

    Private Sub btnCashDepositReport_Click(sender As Object, e As EventArgs) Handles btnCashDepositReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

    End Sub

    Private Sub btnCreditDepositReport_Click(sender As Object, e As EventArgs) Handles btnCreditDepositReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

    End Sub

    Friend WithEvents btnExit As Button

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' close the form
        Me.Close()

    End Sub

    Friend WithEvents Label1 As Label
End Class
