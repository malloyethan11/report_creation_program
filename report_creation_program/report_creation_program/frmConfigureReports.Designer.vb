<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfigureReports
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboReportTypes = New System.Windows.Forms.ComboBox()
        Me.chklstFrequency = New System.Windows.Forms.CheckedListBox()
        Me.btnAccept = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtHour = New System.Windows.Forms.TextBox()
        Me.txtMinute = New System.Windows.Forms.TextBox()
        Me.rdoPM = New System.Windows.Forms.RadioButton()
        Me.rdoAM = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnViewConfig = New System.Windows.Forms.Button()
        Me.dtmPickDate = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Report Type:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 130)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(162, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Frequency (Check all that apply):"
        '
        'cboReportTypes
        '
        Me.cboReportTypes.FormattingEnabled = True
        Me.cboReportTypes.Items.AddRange(New Object() {"Sales Report", "Sales Tax Report", "Inventory Report", "Cash Deposit Report", "Credit Deposit Report"})
        Me.cboReportTypes.Location = New System.Drawing.Point(87, 6)
        Me.cboReportTypes.Name = "cboReportTypes"
        Me.cboReportTypes.Size = New System.Drawing.Size(173, 21)
        Me.cboReportTypes.TabIndex = 2
        '
        'chklstFrequency
        '
        Me.chklstFrequency.FormattingEnabled = True
        Me.chklstFrequency.Items.AddRange(New Object() {"Daily", "Weekly", "Monthly", "Yearly "})
        Me.chklstFrequency.Location = New System.Drawing.Point(180, 130)
        Me.chklstFrequency.Name = "chklstFrequency"
        Me.chklstFrequency.Size = New System.Drawing.Size(80, 64)
        Me.chklstFrequency.TabIndex = 3
        '
        'btnAccept
        '
        Me.btnAccept.Location = New System.Drawing.Point(56, 272)
        Me.btnAccept.Name = "btnAccept"
        Me.btnAccept.Size = New System.Drawing.Size(164, 31)
        Me.btnAccept.TabIndex = 4
        Me.btnAccept.Text = "Accept Configuration"
        Me.btnAccept.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(56, 389)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(164, 31)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 206)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(98, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Time report will run:"
        '
        'txtHour
        '
        Me.txtHour.Location = New System.Drawing.Point(42, 225)
        Me.txtHour.MaxLength = 2
        Me.txtHour.Name = "txtHour"
        Me.txtHour.Size = New System.Drawing.Size(27, 20)
        Me.txtHour.TabIndex = 8
        '
        'txtMinute
        '
        Me.txtMinute.Location = New System.Drawing.Point(102, 225)
        Me.txtMinute.MaxLength = 2
        Me.txtMinute.Name = "txtMinute"
        Me.txtMinute.Size = New System.Drawing.Size(27, 20)
        Me.txtMinute.TabIndex = 9
        '
        'rdoPM
        '
        Me.rdoPM.AutoSize = True
        Me.rdoPM.Location = New System.Drawing.Point(182, 228)
        Me.rdoPM.Name = "rdoPM"
        Me.rdoPM.Size = New System.Drawing.Size(41, 17)
        Me.rdoPM.TabIndex = 10
        Me.rdoPM.TabStop = True
        Me.rdoPM.Text = "PM"
        Me.rdoPM.UseVisualStyleBackColor = True
        '
        'rdoAM
        '
        Me.rdoAM.AutoSize = True
        Me.rdoAM.Location = New System.Drawing.Point(140, 228)
        Me.rdoAM.Name = "rdoAM"
        Me.rdoAM.Size = New System.Drawing.Size(41, 17)
        Me.rdoAM.TabIndex = 11
        Me.rdoAM.TabStop = True
        Me.rdoAM.Text = "AM"
        Me.rdoAM.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 229)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Hour"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(77, 229)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Min"
        '
        'btnViewConfig
        '
        Me.btnViewConfig.Location = New System.Drawing.Point(56, 312)
        Me.btnViewConfig.Name = "btnViewConfig"
        Me.btnViewConfig.Size = New System.Drawing.Size(164, 31)
        Me.btnViewConfig.TabIndex = 14
        Me.btnViewConfig.Text = "View Current Configuration"
        Me.btnViewConfig.UseVisualStyleBackColor = True
        '
        'dtmPickDate
        '
        Me.dtmPickDate.Location = New System.Drawing.Point(60, 80)
        Me.dtmPickDate.Name = "dtmPickDate"
        Me.dtmPickDate.Size = New System.Drawing.Size(200, 20)
        Me.dtmPickDate.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(92, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Start Running On:"
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(56, 350)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(164, 31)
        Me.btnUpdate.TabIndex = 17
        Me.btnUpdate.Text = "Update Configuration"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'frmConfigureReports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(279, 432)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dtmPickDate)
        Me.Controls.Add(Me.btnViewConfig)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.rdoAM)
        Me.Controls.Add(Me.rdoPM)
        Me.Controls.Add(Me.txtMinute)
        Me.Controls.Add(Me.txtHour)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnAccept)
        Me.Controls.Add(Me.chklstFrequency)
        Me.Controls.Add(Me.cboReportTypes)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmConfigureReports"
        Me.Text = "Configure Reports"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents cboReportTypes As ComboBox
    Friend WithEvents chklstFrequency As CheckedListBox
    Friend WithEvents btnAccept As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents txtHour As TextBox
    Friend WithEvents txtMinute As TextBox
    Friend WithEvents rdoPM As RadioButton
    Friend WithEvents rdoAM As RadioButton
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents btnViewConfig As Button
    Friend WithEvents dtmPickDate As DateTimePicker
    Friend WithEvents Label6 As Label
    Friend WithEvents btnUpdate As Button
End Class
