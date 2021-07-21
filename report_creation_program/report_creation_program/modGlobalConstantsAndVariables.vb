Module modGlobalConstantsAndVariables

    Public strConnectionUsername As String
    Public strConnectionPassword As String
    Public btmButtonDefault As Bitmap = My.Resources.Button
    Public btmButtonDefaultGray As Bitmap = My.Resources.Button
    Public btmButtonShort As Bitmap = My.Resources.ButtonShort
    Public btmButtonShortGray As Bitmap = My.Resources.ButtonShort
    ' PLEASE NOTE THAT IN VB.NET THE AXES OR SWAPPED. (Y, X)
    Public aastrCSVFile(,) As String = New String(15, 1) {
        {"Sales Report Daily", "false"},
        {"Sales Report Weekly", "false"},
        {"Sales Report Monthly", "false"},
        {"Sales Report Yearly", "false"},
        {"Sales Tax Report Daily", "false"},
        {"Sales Tax Report Weekly", "false"},
        {"Sales Tax Report Monthly", "false"},
        {"Sales Tax Report Yearly", "false"},
        {"Sales Inventory Daily", "false"},
        {"Sales Inventory Weekly", "false"},
        {"Sales Inventory Monthly", "false"},
        {"Sales Inventory Yearly", "false"},
        {"Sales Deposit Daily", "false"},
        {"Sales Deposit Weekly", "false"},
        {"Sales Deposit Monthly", "false"},
        {"Sales Deposit Yearly", "false"}
    }


    Public strEmailSalesReport As String
    Public strEmailSalesTax As String
    Public strEmailInventory As String
    Public strEmailDesposit As String

    Public blnSalesReportDaily As Boolean
    Public blnSalesReportMonthly As Boolean
    Public blnSalesReportWeekly As Boolean
    Public blnSalesReportYearly As Boolean

    Public blnInventoryReportDaily As Boolean
    Public blnInventoryReportMonthly As Boolean
    Public blnInventoryReportWeekly As Boolean
    Public blnInventoryReportYearly As Boolean

    Public blnSalesTaxReportDaily As Boolean
    Public blnSalesTaxReportMonthly As Boolean
    Public blnSalesTaxReportWeekly As Boolean
    Public blnSalesTaxReportYearly As Boolean

    Public blnDepositReportDaily As Boolean
    Public blnDepositReportMonthly As Boolean
    Public blnDepositReportWeekly As Boolean
    Public blnDepositReportYearly As Boolean

    Public dtmDailySalesReport As DateTime
    Public dtmDailySalesTaxReport As DateTime
    Public dtmDailyInventory As DateTime
    Public dtmDailyDeposit As DateTime

    Public dtmWeeklySalesReport As DateTime
    Public dtmWeeklySalesTaxReport As DateTime
    Public dtmWeeklyInventory As DateTime
    Public dtmWeeklyDeposit As DateTime

    Public dtmMonthlySalesReport As DateTime
    Public dtmMonthlySalesTaxReport As DateTime
    Public dtmMonthlyInventory As DateTime
    Public dtmMonthlyDeposit As DateTime

    Public dtmYearlySalesReport As DateTime
    Public dtmYearlySalesTaxReport As DateTime
    Public dtmYearlyInventory As DateTime
    Public dtmYearlyDeposit As DateTime

End Module
