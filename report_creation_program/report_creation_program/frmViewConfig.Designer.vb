<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewConfig
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblReportType = New System.Windows.Forms.Label()
        Me.lblDaily = New System.Windows.Forms.Label()
        Me.lblWeekly = New System.Windows.Forms.Label()
        Me.lblMonthly = New System.Windows.Forms.Label()
        Me.lblYearly = New System.Windows.Forms.Label()
        Me.lblDayofWeek = New System.Windows.Forms.Label()
        Me.lblTime = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Report Type:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(23, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Run Daily?"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 66)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Run Weekly?"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(23, 91)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Run Monthly?"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(23, 114)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Run Yearly?"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(23, 138)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(114, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Day of the Week Run:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(23, 162)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(90, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Time of Day Run:"
        '
        'lblReportType
        '
        Me.lblReportType.AutoSize = True
        Me.lblReportType.Location = New System.Drawing.Point(150, 20)
        Me.lblReportType.Name = "lblReportType"
        Me.lblReportType.Size = New System.Drawing.Size(39, 13)
        Me.lblReportType.TabIndex = 7
        Me.lblReportType.Text = "Label8"
        '
        'lblDaily
        '
        Me.lblDaily.AutoSize = True
        Me.lblDaily.Location = New System.Drawing.Point(150, 43)
        Me.lblDaily.Name = "lblDaily"
        Me.lblDaily.Size = New System.Drawing.Size(39, 13)
        Me.lblDaily.TabIndex = 8
        Me.lblDaily.Text = "Label9"
        '
        'lblWeekly
        '
        Me.lblWeekly.AutoSize = True
        Me.lblWeekly.Location = New System.Drawing.Point(150, 66)
        Me.lblWeekly.Name = "lblWeekly"
        Me.lblWeekly.Size = New System.Drawing.Size(45, 13)
        Me.lblWeekly.TabIndex = 9
        Me.lblWeekly.Text = "Label10"
        '
        'lblMonthly
        '
        Me.lblMonthly.AutoSize = True
        Me.lblMonthly.Location = New System.Drawing.Point(150, 91)
        Me.lblMonthly.Name = "lblMonthly"
        Me.lblMonthly.Size = New System.Drawing.Size(45, 13)
        Me.lblMonthly.TabIndex = 10
        Me.lblMonthly.Text = "Label11"
        '
        'lblYearly
        '
        Me.lblYearly.AutoSize = True
        Me.lblYearly.Location = New System.Drawing.Point(150, 114)
        Me.lblYearly.Name = "lblYearly"
        Me.lblYearly.Size = New System.Drawing.Size(45, 13)
        Me.lblYearly.TabIndex = 11
        Me.lblYearly.Text = "Label12"
        '
        'lblDayofWeek
        '
        Me.lblDayofWeek.AutoSize = True
        Me.lblDayofWeek.Location = New System.Drawing.Point(150, 138)
        Me.lblDayofWeek.Name = "lblDayofWeek"
        Me.lblDayofWeek.Size = New System.Drawing.Size(45, 13)
        Me.lblDayofWeek.TabIndex = 12
        Me.lblDayofWeek.Text = "Label13"
        '
        'lblTime
        '
        Me.lblTime.AutoSize = True
        Me.lblTime.Location = New System.Drawing.Point(150, 162)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(45, 13)
        Me.lblTime.TabIndex = 13
        Me.lblTime.Text = "Label14"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(26, 192)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(95, 32)
        Me.btnOK.TabIndex = 14
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'frmViewConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(301, 236)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lblTime)
        Me.Controls.Add(Me.lblDayofWeek)
        Me.Controls.Add(Me.lblYearly)
        Me.Controls.Add(Me.lblMonthly)
        Me.Controls.Add(Me.lblWeekly)
        Me.Controls.Add(Me.lblDaily)
        Me.Controls.Add(Me.lblReportType)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmViewConfig"
        Me.Text = "View Report Configuration"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents lblReportType As Label
    Friend WithEvents lblDaily As Label
    Friend WithEvents lblWeekly As Label
    Friend WithEvents lblMonthly As Label
    Friend WithEvents lblYearly As Label
    Friend WithEvents lblDayofWeek As Label
    Friend WithEvents lblTime As Label
    Friend WithEvents btnOK As Button
End Class
