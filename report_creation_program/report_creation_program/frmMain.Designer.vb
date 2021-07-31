Imports Excel = Microsoft.Office.Interop.Excel

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

    Friend WithEvents Label1 As Label


    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnCashCreditDepositReport = New System.Windows.Forms.Button()
        Me.btnInventoryReport = New System.Windows.Forms.Button()
        Me.btnTaxReport = New System.Windows.Forms.Button()
        Me.btnSalesReport = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tmrUpdateLocalConfig = New System.Windows.Forms.Timer(Me.components)
        Me.tmrCheckIfReportRun = New System.Windows.Forms.Timer(Me.components)
        Me.tmrUpdateButtonImage = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'btnCashCreditDepositReport
        '
        Me.btnCashCreditDepositReport.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnCashCreditDepositReport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCashCreditDepositReport.Image = CType(resources.GetObject("btnCashCreditDepositReport.Image"), System.Drawing.Image)
        Me.btnCashCreditDepositReport.Location = New System.Drawing.Point(19, 156)
        Me.btnCashCreditDepositReport.Name = "btnCashCreditDepositReport"
        Me.btnCashCreditDepositReport.Size = New System.Drawing.Size(112, 41)
        Me.btnCashCreditDepositReport.TabIndex = 1
        Me.btnCashCreditDepositReport.Text = "Deposit Report"
        Me.btnCashCreditDepositReport.UseVisualStyleBackColor = True
        '
        'btnInventoryReport
        '
        Me.btnInventoryReport.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnInventoryReport.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInventoryReport.Image = CType(resources.GetObject("btnInventoryReport.Image"), System.Drawing.Image)
        Me.btnInventoryReport.Location = New System.Drawing.Point(19, 109)
        Me.btnInventoryReport.Name = "btnInventoryReport"
        Me.btnInventoryReport.Size = New System.Drawing.Size(112, 41)
        Me.btnInventoryReport.TabIndex = 2
        Me.btnInventoryReport.Text = "Inventory Report"
        Me.btnInventoryReport.UseVisualStyleBackColor = True
        '
        'btnTaxReport
        '
        Me.btnTaxReport.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnTaxReport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTaxReport.Image = CType(resources.GetObject("btnTaxReport.Image"), System.Drawing.Image)
        Me.btnTaxReport.Location = New System.Drawing.Point(19, 62)
        Me.btnTaxReport.Name = "btnTaxReport"
        Me.btnTaxReport.Size = New System.Drawing.Size(112, 41)
        Me.btnTaxReport.TabIndex = 3
        Me.btnTaxReport.Text = "Tax Report"
        Me.btnTaxReport.UseVisualStyleBackColor = True
        '
        'btnSalesReport
        '
        Me.btnSalesReport.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnSalesReport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSalesReport.Image = Global.TestNET.My.Resources.Resources.ButtonShort
        Me.btnSalesReport.Location = New System.Drawing.Point(19, 15)
        Me.btnSalesReport.Name = "btnSalesReport"
        Me.btnSalesReport.Size = New System.Drawing.Size(112, 41)
        Me.btnSalesReport.TabIndex = 4
        Me.btnSalesReport.Text = "Sales Report"
        Me.btnSalesReport.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnExit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = Global.TestNET.My.Resources.Resources.ButtonShort
        Me.btnExit.Location = New System.Drawing.Point(19, 235)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(112, 41)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 204)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(139, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "______________________"
        '
        'tmrUpdateLocalConfig
        '
        Me.tmrUpdateLocalConfig.Enabled = True
        Me.tmrUpdateLocalConfig.Interval = 1000
        '
        'tmrCheckIfReportRun
        '
        Me.tmrCheckIfReportRun.Enabled = True
        Me.tmrCheckIfReportRun.Interval = 1000
        '
        'tmrUpdateButtonImage
        '
        Me.tmrUpdateButtonImage.Enabled = True
        Me.tmrUpdateButtonImage.Interval = 1
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(151, 290)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSalesReport)
        Me.Controls.Add(Me.btnTaxReport)
        Me.Controls.Add(Me.btnInventoryReport)
        Me.Controls.Add(Me.btnCashCreditDepositReport)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "Home"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCashCreditDepositReport As Button
    Friend WithEvents btnInventoryReport As Button
    Friend WithEvents btnTaxReport As Button
    Friend WithEvents btnSalesReport As Button

    Friend WithEvents tmrUpdateLocalConfig As Timer
    Friend WithEvents tmrCheckIfReportRun As Timer
    Friend WithEvents tmrUpdateButtonImage As Timer
End Class
