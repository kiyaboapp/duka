Attribute VB_Name = "modExpenses"
Option Compare Database
Option Explicit
'modExpenses

' Get total expenses for given criteria with date range support
Public Function getExpenses(Optional expenseItemID As Long = 0, _
                           Optional expenseTypeID As Long = 0, _
                           Optional startDate As Variant = "", _
                           Optional EndDate As Variant = "") As Currency
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim dtStart As Date
    Dim dtEnd As Date
    
    Set db = CurrentDb
    getExpenses = 0
    
    ' Determine date range
    If startDate = "" And EndDate = "" Then
        ' No dates specified - use all records up to today
        dtEnd = Date
        sql = "SELECT SUM(o.amount_due) AS TotalExpense " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    ElseIf startDate = "" And EndDate <> "" Then
        ' Only end date specified - all records up to end date
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.amount_due) AS TotalExpense " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    ElseIf startDate <> "" And EndDate = "" Then
        ' Only start date specified - from start date to today
        dtStart = DateValue(CDate(startDate))
        dtEnd = Date
        sql = "SELECT SUM(o.amount_due) AS TotalExpense " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    Else
        ' Both dates specified - use date range
        dtStart = DateValue(CDate(startDate))
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.amount_due) AS TotalExpense " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    End If
    
    ' Add optional filters
    If expenseItemID > 0 Then
        sql = sql & "AND o.expense_item_id = " & expenseItemID & " "
    End If
    
    If expenseTypeID > 0 Then
        sql = sql & "AND ei.expense_type_id = " & expenseTypeID & " "
    End If
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        If Not IsNull(rs!totalExpense) Then
            getExpenses = rs!totalExpense
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    getExpenses = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Get total paid amount for given criteria with date range support
Public Function GetPaidAmount(Optional expenseItemID As Long = 0, _
                             Optional expenseTypeID As Long = 0, _
                             Optional startDate As Variant = "", _
                             Optional EndDate As Variant = "") As Currency
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim dtStart As Date
    Dim dtEnd As Date
    
    Set db = CurrentDb
    GetPaidAmount = 0
    
    ' Determine date range
    If startDate = "" And EndDate = "" Then
        dtEnd = Date
        sql = "SELECT SUM(o.amount_paid) AS TotalPaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    ElseIf startDate = "" And EndDate <> "" Then
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.amount_paid) AS TotalPaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    ElseIf startDate <> "" And EndDate = "" Then
        dtStart = DateValue(CDate(startDate))
        dtEnd = Date
        sql = "SELECT SUM(o.amount_paid) AS TotalPaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    Else
        dtStart = DateValue(CDate(startDate))
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.amount_paid) AS TotalPaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    End If
    
    ' Add optional filters
    If expenseItemID > 0 Then
        sql = sql & "AND o.expense_item_id = " & expenseItemID & " "
    End If
    
    If expenseTypeID > 0 Then
        sql = sql & "AND ei.expense_type_id = " & expenseTypeID & " "
    End If
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        If Not IsNull(rs!totalPaid) Then
            GetPaidAmount = rs!totalPaid
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetPaidAmount = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Get total prepaid amount for given criteria with date range support
Public Function GetPrepaidAmount(Optional expenseItemID As Long = 0, _
                                Optional expenseTypeID As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "") As Currency
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim dtStart As Date
    Dim dtEnd As Date
    
    Set db = CurrentDb
    GetPrepaidAmount = 0
    
    ' Determine date range
    If startDate = "" And EndDate = "" Then
        dtEnd = Date
        sql = "SELECT SUM(o.prepayment_applied) AS TotalPrepaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    ElseIf startDate = "" And EndDate <> "" Then
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.prepayment_applied) AS TotalPrepaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    ElseIf startDate <> "" And EndDate = "" Then
        dtStart = DateValue(CDate(startDate))
        dtEnd = Date
        sql = "SELECT SUM(o.prepayment_applied) AS TotalPrepaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    Else
        dtStart = DateValue(CDate(startDate))
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.prepayment_applied) AS TotalPrepaid " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# "
    End If
    
    ' Add optional filters
    If expenseItemID > 0 Then
        sql = sql & "AND o.expense_item_id = " & expenseItemID & " "
    End If
    
    If expenseTypeID > 0 Then
        sql = sql & "AND ei.expense_type_id = " & expenseTypeID & " "
    End If
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        If Not IsNull(rs!totalPrepaid) Then
            GetPrepaidAmount = rs!totalPrepaid
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetPrepaidAmount = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Get total pending/outstanding balance for given criteria with date range support
Public Function GetPendingAmount(Optional expenseItemID As Long = 0, _
                                Optional expenseTypeID As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "") As Currency
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim dtStart As Date
    Dim dtEnd As Date
    
    Set db = CurrentDb
    GetPendingAmount = 0
    
    ' Determine date range
    If startDate = "" And EndDate = "" Then
        dtEnd = Date
        sql = "SELECT SUM(o.balance) AS TotalPending " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# " & _
              "AND o.balance > 0 "
    ElseIf startDate = "" And EndDate <> "" Then
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.balance) AS TotalPending " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# " & _
              "AND o.balance > 0 "
    ElseIf startDate <> "" And EndDate = "" Then
        dtStart = DateValue(CDate(startDate))
        dtEnd = Date
        sql = "SELECT SUM(o.balance) AS TotalPending " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# " & _
              "AND o.balance > 0 "
    Else
        dtStart = DateValue(CDate(startDate))
        dtEnd = DateValue(CDate(EndDate))
        sql = "SELECT SUM(o.balance) AS TotalPending " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) <= #" & Format(dtEnd, "mm/dd/yyyy") & "# " & _
              "AND o.balance > 0 "
    End If
    
    ' Add optional filters
    If expenseItemID > 0 Then
        sql = sql & "AND o.expense_item_id = " & expenseItemID & " "
    End If
    
    If expenseTypeID > 0 Then
        sql = sql & "AND ei.expense_type_id = " & expenseTypeID & " "
    End If
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        If Not IsNull(rs!totalPending) Then
            GetPendingAmount = rs!totalPending
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetPendingAmount = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Get available prepayment balance for given criteria (no date range needed)
Public Function GetAvailablePrepaymentBalance(Optional expenseItemID As Long = 0, _
                                             Optional expenseTypeID As Long = 0) As Currency
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    GetAvailablePrepaymentBalance = 0
    
    ' Build SQL - sum of remaining prepayments
    sql = "SELECT SUM(p.total_prepaid - p.amount_utilized) AS AvailableBalance " & _
          "FROM tbl_prepayments p " & _
          "INNER JOIN tbl_expense_items ei ON p.expense_item_id = ei.expense_item_id " & _
          "WHERE (p.total_prepaid - p.amount_utilized) > 0 "
    
    ' Add optional filters
    If expenseItemID > 0 Then
        sql = sql & "AND p.expense_item_id = " & expenseItemID & " "
    End If
    
    If expenseTypeID > 0 Then
        sql = sql & "AND ei.expense_type_id = " & expenseTypeID & " "
    End If
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        If Not IsNull(rs!AvailableBalance) Then
            GetAvailablePrepaymentBalance = rs!AvailableBalance
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetAvailablePrepaymentBalance = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Get overdue amount (past due date and still unpaid) with date range support
Public Function GetOverdueAmount(Optional expenseItemID As Long = 0, _
                                Optional expenseTypeID As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "") As Currency
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim dtStart As Date
    Dim dtEnd As Date
    
    Set db = CurrentDb
    GetOverdueAmount = 0
    
    ' Determine date range - for overdue, we check obligations before the end date
    If EndDate = "" Then
        dtEnd = Date
    Else
        dtEnd = DateValue(CDate(EndDate))
    End If
    
    ' Build SQL - obligations past due date with balance > 0
    If startDate = "" Then
        ' No start date - all overdue items before end date
        sql = "SELECT SUM(o.balance) AS TotalOverdue " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) < #" & Format(dtEnd, "mm/dd/yyyy") & "# " & _
              "AND o.balance > 0 "
    Else
        ' Start date specified - overdue items within date range
        dtStart = DateValue(CDate(startDate))
        sql = "SELECT SUM(o.balance) AS TotalOverdue " & _
              "FROM tbl_payment_obligations o " & _
              "INNER JOIN tbl_expense_items ei ON o.expense_item_id = ei.expense_item_id " & _
              "WHERE DateValue(o.due_date) >= #" & Format(dtStart, "mm/dd/yyyy") & "# " & _
              "AND DateValue(o.due_date) < #" & Format(dtEnd, "mm/dd/yyyy") & "# " & _
              "AND o.balance > 0 "
    End If
    
    ' Add optional filters
    If expenseItemID > 0 Then
        sql = sql & "AND o.expense_item_id = " & expenseItemID & " "
    End If
    
    If expenseTypeID > 0 Then
        sql = sql & "AND ei.expense_type_id = " & expenseTypeID & " "
    End If
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        If Not IsNull(rs!totalOverdue) Then
            GetOverdueAmount = rs!totalOverdue
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetOverdueAmount = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Get expense summary report with date range support
Public Sub PrintExpenseSummary(Optional expenseItemID As Long = 0, _
                              Optional expenseTypeID As Long = 0, _
                              Optional startDate As Variant = "", _
                              Optional EndDate As Variant = "")
    
    Dim totalExpense As Currency
    Dim totalPaid As Currency
    Dim totalPrepaid As Currency
    Dim totalPending As Currency
    Dim totalOverdue As Currency
    Dim availablePrepayment As Currency
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim reportTitle As String
    Dim dateRangeText As String
    
    ' Determine date range for display
    If startDate = "" And EndDate = "" Then
        dtEnd = Date
        dateRangeText = "Up to " & Format(dtEnd, "mm/dd/yyyy")
    ElseIf startDate = "" And EndDate <> "" Then
        dtEnd = DateValue(CDate(EndDate))
        dateRangeText = "Up to " & Format(dtEnd, "mm/dd/yyyy")
    ElseIf startDate <> "" And EndDate = "" Then
        dtStart = DateValue(CDate(startDate))
        dtEnd = Date
        dateRangeText = Format(dtStart, "mm/dd/yyyy") & " to " & Format(dtEnd, "mm/dd/yyyy")
    Else
        dtStart = DateValue(CDate(startDate))
        dtEnd = DateValue(CDate(EndDate))
        dateRangeText = Format(dtStart, "mm/dd/yyyy") & " to " & Format(dtEnd, "mm/dd/yyyy")
    End If
    
    ' Get all amounts
    totalExpense = getExpenses(expenseItemID, expenseTypeID, startDate, EndDate)
    totalPaid = GetPaidAmount(expenseItemID, expenseTypeID, startDate, EndDate)
    totalPrepaid = GetPrepaidAmount(expenseItemID, expenseTypeID, startDate, EndDate)
    totalPending = GetPendingAmount(expenseItemID, expenseTypeID, startDate, EndDate)
    totalOverdue = GetOverdueAmount(expenseItemID, expenseTypeID, startDate, EndDate)
    availablePrepayment = GetAvailablePrepaymentBalance(expenseItemID, expenseTypeID)
    
    ' Build report title
    reportTitle = "EXPENSE SUMMARY REPORT"
    If expenseItemID > 0 Then
        reportTitle = reportTitle & " - Item ID: " & expenseItemID
    End If
    If expenseTypeID > 0 Then
        reportTitle = reportTitle & " - Type ID: " & expenseTypeID
    End If
    reportTitle = reportTitle & " (" & dateRangeText & ")"
    
    ' Print report to Debug window
    Debug.Print String(80, "=")
    Debug.Print reportTitle
    Debug.Print String(80, "=")
    Debug.Print "Date Range: " & dateRangeText
    Debug.Print String(80, "-")
    Debug.Print "Total Expense (Due): " & Format(totalExpense, "Currency")
    Debug.Print "Total Paid: " & Format(totalPaid, "Currency")
    Debug.Print "Total Prepaid Applied: " & Format(totalPrepaid, "Currency")
    Debug.Print String(80, "-")
    Debug.Print "Total Covered: " & Format(totalPaid + totalPrepaid, "Currency")
    Debug.Print "Pending Balance: " & Format(totalPending, "Currency")
    Debug.Print "Overdue Amount: " & Format(totalOverdue, "Currency")
    Debug.Print String(80, "-")
    Debug.Print "Available Prepayments: " & Format(availablePrepayment, "Currency")
    Debug.Print String(80, "=")
    
    ' Show in message box
    MsgBox "Date Range: " & dateRangeText & vbCrLf & vbCrLf & _
           "Total Expense: " & Format(totalExpense, "Currency") & vbCrLf & _
           "Total Paid: " & Format(totalPaid, "Currency") & vbCrLf & _
           "Total Prepaid: " & Format(totalPrepaid, "Currency") & vbCrLf & _
           "Pending Balance: " & Format(totalPending, "Currency") & vbCrLf & _
           "Overdue: " & Format(totalOverdue, "Currency") & vbCrLf & _
           "Available Prepayments: " & Format(availablePrepayment, "Currency"), _
           vbInformation, reportTitle
           
End Sub

