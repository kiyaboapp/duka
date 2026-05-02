Attribute VB_Name = "modPayables"
Option Compare Database
Option Explicit

'==================================================================================================
' modPayables - FINANCIAL MANAGEMENT SYSTEM
' Supports: Expenses, Liabilities, Obligations, Prepayments, Payments
' Fixed for Calculated "balance" field - NEVER writes to balance
' Enhanced: Historical generation, last day + negative offsets, debug support
'==================================================================================================

' ================================================================================
' HELPER: Smart Due Date Calculation (supports last day and negative offsets)
' ================================================================================
Private Function CalculateDueDate(recurrenceType As String, _
                                  baseDate As Date, _
                                  specificDayOfMonth As Variant, _
                                  specificDayOfWeek As Variant) As Date
    Dim lastDayOfMonth As Date
    Dim offset As Long
    
    On Error GoTo ErrorHandler
    
    Select Case UCase(recurrenceType)
        Case "MONTHLY"
            lastDayOfMonth = DateSerial(Year(baseDate), Month(baseDate) + 1, 0)
            
            If IsNull(specificDayOfMonth) Or specificDayOfMonth = 0 Then
                offset = 0
            ElseIf specificDayOfMonth < 0 Then
                offset = specificDayOfMonth
            Else
                CalculateDueDate = DateSerial(Year(baseDate), Month(baseDate), specificDayOfMonth)
                Exit Function
            End If
            CalculateDueDate = DateAdd("d", offset, lastDayOfMonth)
            
        Case "WEEKLY"
            If IsNull(specificDayOfWeek) Then specificDayOfWeek = 1
            
            If specificDayOfWeek = 0 Then
                CalculateDueDate = baseDate - Weekday(baseDate, vbMonday) + 7
            ElseIf specificDayOfWeek < 0 Then
                CalculateDueDate = (baseDate - Weekday(baseDate, vbMonday) + 7) + specificDayOfWeek
            Else
                CalculateDueDate = baseDate - Weekday(baseDate, vbMonday) + specificDayOfWeek
            End If
            
        Case Else
            CalculateDueDate = baseDate
    End Select
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in CalculateDueDate: " & Err.description
    CalculateDueDate = baseDate
End Function

' ================================================================================
' SECTION 1: EXPENSE ITEM MANAGEMENT
' ================================================================================
Public Function CreateExpenseItem(expenseTypeID As Long, _
                                  itemName As String, _
                                  itemDescription As String, _
                                  startDate As Date, _
                                  Optional EndDate As Variant = Null) As Long
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim newID As Long
   
    Set db = CurrentDb
    CreateExpenseItem = 0
   
    If Len(Trim(itemName)) = 0 Then
        MsgBox "Item name cannot be empty!", vbCritical, "Validation Error"
        Exit Function
    End If
   
    Set rs = db.OpenRecordset("tbl_expense_items", dbOpenDynaset)
    rs.AddNew
    rs!expense_type_id = expenseTypeID
    rs!item_name = Trim(itemName)
    rs!item_description = Trim(itemDescription)
    rs!start_date = startDate
    If IsDate(EndDate) Then rs!end_date = CDate(EndDate)
    rs!is_active = True
    rs.upDate
   
    rs.Bookmark = rs.LastModified
    newID = rs!expense_item_id
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    CreateExpenseItem = newID
    MsgBox "Expense item '" & itemName & "' created successfully!" & vbCrLf & _
           "Expense Item ID: " & newID, vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error creating expense item: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    CreateExpenseItem = 0
End Function

Public Function UpdateExpenseItem(expenseItemID As Long, _
                                  Optional itemName As Variant = Null, _
                                  Optional itemDescription As Variant = Null, _
                                  Optional EndDate As Variant = Null, _
                                  Optional isActive As Variant = Null) As Boolean
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
   
    Set db = CurrentDb
    UpdateExpenseItem = False
   
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tbl_expense_items WHERE expense_item_id = " & expenseItemID, _
        dbOpenDynaset)
   
    If rs.EOF Then
        MsgBox "Expense item ID " & expenseItemID & " not found!", vbCritical, "Error"
        rs.Close
        Set db = Nothing
        Exit Function
    End If
   
    rs.Edit
    If Not IsNull(itemName) Then rs!item_name = Trim(itemName)
    If Not IsNull(itemDescription) Then rs!item_description = Trim(itemDescription)
    If IsDate(EndDate) Then rs!end_date = CDate(EndDate)
    If Not IsNull(isActive) Then rs!is_active = CBool(isActive)
    rs.upDate
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    UpdateExpenseItem = True
    MsgBox "Expense item updated successfully!", vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error updating expense item: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    UpdateExpenseItem = False
End Function

Public Function DeleteExpenseItem(expenseItemID As Long, Optional hardDelete As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
   
    Set db = CurrentDb
    DeleteExpenseItem = False
   
    If hardDelete Then
        Set rs = db.OpenRecordset( _
            "SELECT COUNT(*) AS cnt FROM tbl_payment_obligations WHERE expense_item_id = " & expenseItemID)
        If Not rs.EOF Then
            If rs!cnt > 0 Then
                MsgBox "Cannot delete: This item has " & rs!cnt & " related obligations!", _
                       vbCritical, "Delete Failed"
                rs.Close
                Set db = Nothing
                Exit Function
            End If
        End If
        rs.Close
       
        db.Execute "DELETE FROM tbl_expense_rates WHERE expense_item_id = " & expenseItemID
        db.Execute "DELETE FROM tbl_recurrence_patterns WHERE expense_item_id = " & expenseItemID
        db.Execute "DELETE FROM tbl_expense_items WHERE expense_item_id = " & expenseItemID
        MsgBox "Expense item permanently deleted!", vbInformation, "Success"
    Else
        Set rs = db.OpenRecordset( _
            "SELECT * FROM tbl_expense_items WHERE expense_item_id = " & expenseItemID, _
            dbOpenDynaset)
       
        If Not rs.EOF Then
            rs.Edit
            rs!is_active = False
            rs!end_date = Date
            rs.upDate
            MsgBox "Expense item deactivated!", vbInformation, "Success"
        End If
        rs.Close
    End If
   
    Set rs = Nothing
    Set db = Nothing
    DeleteExpenseItem = True
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error deleting expense item: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    DeleteExpenseItem = False
End Function

' ================================================================================
' SECTION 2: EXPENSE RATE MANAGEMENT
' ================================================================================
Public Function AddExpenseRate(expenseItemID As Long, _
                              newRate As Currency, _
                              effectiveDate As Date, _
                              changeReason As String) As Long
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rsCheck As DAO.Recordset
    Dim rsNewRate As DAO.Recordset
    Dim newRateID As Long
   
    Set db = CurrentDb
    AddExpenseRate = 0
   
    Set rsCheck = db.OpenRecordset( _
        "SELECT expense_item_id FROM tbl_expense_items WHERE expense_item_id = " & expenseItemID)
   
    If rsCheck.EOF Then
        MsgBox "Expense item ID " & expenseItemID & " not found!", vbCritical, "Error"
        rsCheck.Close
        Set db = Nothing
        Exit Function
    End If
    rsCheck.Close
   
    If newRate <= 0 Then
        MsgBox "New rate must be greater than zero!", vbCritical, "Error"
        Set db = Nothing
        Exit Function
    End If
   
    Set rsNewRate = db.OpenRecordset("tbl_expense_rates", dbOpenDynaset)
    rsNewRate.AddNew
    rsNewRate!expense_item_id = expenseItemID
    rsNewRate!rate_amount = newRate
    rsNewRate!effective_from = effectiveDate
    If Len(changeReason) > 0 Then rsNewRate!change_reason = changeReason
    rsNewRate.upDate
   
    rsNewRate.Bookmark = rsNewRate.LastModified
    newRateID = rsNewRate!rate_id
    rsNewRate.Close
    Set rsNewRate = Nothing
    Set db = Nothing
   
    AddExpenseRate = newRateID
   
    MsgBox "Rate changed to " & Format(newRate, "Currency") & " effective " & _
           Format(effectiveDate, "Short Date") & vbCrLf & _
           "Rate ID: " & newRateID, vbInformation, "Rate Changed"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error changing rate: " & Err.description, vbCritical, "Error"
    On Error Resume Next
    If Not rsCheck Is Nothing Then rsCheck.Close
    If Not rsNewRate Is Nothing Then rsNewRate.Close
    Set rsCheck = Nothing
    Set rsNewRate = Nothing
    Set db = Nothing
    AddExpenseRate = 0
End Function

Public Function GetCurrentRate(expenseItemID As Long, Optional targetDate As Variant = Null) As Currency
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rateValue As Currency
    Dim checkDate As Date
   
    Set db = CurrentDb
    rateValue = 0
   
    If IsDate(targetDate) Then
        checkDate = CDate(targetDate)
    Else
        checkDate = Date
    End If
   
    Set rs = db.OpenRecordset( _
        "SELECT TOP 1 rate_amount FROM tbl_expense_rates " & _
        "WHERE expense_item_id = " & expenseItemID & _
        " AND effective_from <= #" & Format(checkDate, "mm/dd/yyyy") & "# " & _
        "ORDER BY effective_from DESC")
   
    If Not rs.EOF Then
        rateValue = rs!rate_amount
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetCurrentRate = rateValue
End Function

' ================================================================================
' SECTION 3: RECURRENCE PATTERN MANAGEMENT
' ================================================================================
Public Function CreateRecurrencePattern(expenseItemID As Long, _
                                        recurrenceType As String, _
                                        startDate As Date, _
                                        Optional specificDayOfWeek As Variant = Null, _
                                        Optional specificDayOfMonth As Variant = Null, _
                                        Optional EndDate As Variant = Null) As Long
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim newID As Long
   
    Set db = CurrentDb
    CreateRecurrencePattern = 0
   
    Set rs = db.OpenRecordset("tbl_recurrence_patterns", dbOpenDynaset)
    rs.AddNew
    rs!expense_item_id = expenseItemID
    rs!recurrence_type = recurrenceType
    rs!start_date = startDate
    If Not IsNull(specificDayOfWeek) Then rs!specific_day_of_week = specificDayOfWeek
    If Not IsNull(specificDayOfMonth) Then rs!specific_day_of_month = specificDayOfMonth
    If IsDate(EndDate) Then rs!end_date = CDate(EndDate)
    rs!is_active = True
    rs.upDate
   
    rs.Bookmark = rs.LastModified
    newID = rs!recurrence_id
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    CreateRecurrencePattern = newID
    MsgBox "Recurrence pattern created successfully!" & vbCrLf & "Pattern ID: " & newID, _
           vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error creating recurrence pattern: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    CreateRecurrencePattern = 0
End Function

Public Function UpdateRecurrencePattern(recurrenceID As Long, _
                                        Optional isActive As Variant = Null, _
                                        Optional EndDate As Variant = Null) As Boolean
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
   
    Set db = CurrentDb
    UpdateRecurrencePattern = False
   
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tbl_recurrence_patterns WHERE recurrence_id = " & recurrenceID, _
        dbOpenDynaset)
   
    If rs.EOF Then
        MsgBox "Recurrence pattern not found!", vbCritical, "Error"
        rs.Close
        Set db = Nothing
        Exit Function
    End If
   
    rs.Edit
    If Not IsNull(isActive) Then rs!is_active = CBool(isActive)
    If IsDate(EndDate) Then rs!end_date = CDate(EndDate)
    rs.upDate
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    UpdateRecurrencePattern = True
    MsgBox "Recurrence pattern updated!", vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error updating recurrence pattern: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    UpdateRecurrencePattern = False
End Function

' ================================================================================
' SECTION 4: OBLIGATION GENERATION (IMPROVED & SAFE)
' ================================================================================
Public Sub GenerateObligations(recurrenceType As String, Optional targetDate As Variant = Null)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsItems As DAO.Recordset
    Dim rsCheck As DAO.Recordset
    Dim dueDate As Date
    Dim currentRate As Currency
    Dim obligationExists As Boolean
    Dim recordsCreated As Integer
    Dim targetBase As Date
    
    Set db = CurrentDb
    recordsCreated = 0
    
    If IsDate(targetDate) Then
        targetBase = CDate(targetDate)
        Debug.Print "GenerateObligations: Target date = " & Format(targetBase, "yyyy-mm-dd")
    Else
        targetBase = Date
        Debug.Print "GenerateObligations: Using today = " & Format(targetBase, "yyyy-mm-dd")
    End If
    
    Dim sql As String
    sql = "SELECT rp.recurrence_id, ei.expense_item_id, ei.item_name, " & _
          "rp.specific_day_of_week, rp.specific_day_of_month " & _
          "FROM tbl_expense_items ei " & _
          "INNER JOIN tbl_recurrence_patterns rp ON ei.expense_item_id = rp.expense_item_id " & _
          "WHERE rp.recurrence_type = '" & recurrenceType & "' " & _
          "AND rp.is_active = True AND ei.is_active = True"
    
    Set rsItems = db.OpenRecordset(sql)
    
    If rsItems.EOF Then
        Debug.Print "No active " & recurrenceType & " patterns found."
        GoTo CleanExit
    End If
    
    Do While Not rsItems.EOF
        dueDate = CalculateDueDate(recurrenceType, targetBase, _
                                   rsItems!specific_day_of_month, rsItems!specific_day_of_week)
        
        Debug.Print "Item: " & rsItems!item_name & " | Calculated Due: " & Format(dueDate, "yyyy-mm-dd")
        
        Set rsCheck = db.OpenRecordset( _
            "SELECT obligation_id FROM tbl_payment_obligations " & _
            "WHERE expense_item_id = " & rsItems!expense_item_id & _
            " AND due_date = #" & Format(dueDate, "mm/dd/yyyy") & "#")
        
        obligationExists = Not rsCheck.EOF
        rsCheck.Close
        Set rsCheck = Nothing
        
        If obligationExists Then
            Debug.Print "  -> Already exists. Skipping."
        Else
            currentRate = GetCurrentRate(rsItems!expense_item_id, dueDate)
            If currentRate > 0 Then
                Dim rsNew As DAO.Recordset
                Set rsNew = db.OpenRecordset("tbl_payment_obligations", dbOpenDynaset)
                rsNew.AddNew
                rsNew!expense_item_id = rsItems!expense_item_id
                rsNew!obligation_type = "Expense"
                rsNew!obligation_date = Date
                rsNew!due_date = dueDate
                rsNew!amount_due = currentRate
                rsNew!amount_paid = 0
                rsNew!prepayment_applied = 0
                rsNew!payment_status = "Pending"
                rsNew!description = "Auto-generated " & recurrenceType & " obligation"
                rsNew.upDate
                rsNew.Close
                Set rsNew = Nothing
                
                recordsCreated = recordsCreated + 1
                Debug.Print "  -> CREATED: " & Format(currentRate, "Currency") & " due " & dueDate
            Else
                Debug.Print "  -> No valid rate. Skipping."
            End If
        End If
        
        rsItems.MoveNext
    Loop
    
CleanExit:
    If Not rsItems Is Nothing Then rsItems.Close
    Set rsItems = Nothing
    Set db = Nothing
    
    Debug.Print "Created " & recordsCreated & " " & recurrenceType & " obligation(s)", vbInformation, "Generation Complete"
    Debug.Print "GenerateObligations finished. Total created: " & recordsCreated
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating obligations: " & Err.description, vbCritical, "Error"
    Debug.Print "ERROR in GenerateObligations: " & Err.description
    Resume CleanExit
End Sub

' ================================================================================
' SECTION 5: PREPAYMENT RECORDING
' ================================================================================
Public Function RecordPrepayment(itemID As Long, _
                                prepaymentAmount As Currency, _
                                Optional prepaymentDate As Variant = Null, _
                                Optional paymentMethodId As Long = 0, _
                                Optional description As String = "", _
                                Optional isExpense As Boolean = True _
                                ) As Long
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rsCheck As DAO.Recordset
    Dim rsPayment As DAO.Recordset
    Dim rsPrepayment As DAO.Recordset
    Dim newPaymentID As Long
    Dim newPrepaymentID As Long
    Dim tableName As String
    Dim fieldName As String
   
    Set db = CurrentDb
    RecordPrepayment = 0
   
    ' Determine which table and field to check
    If isExpense Then
        tableName = "tbl_expense_items"
        fieldName = "expense_item_id"
    Else
        tableName = "tbl_liability_items"
        fieldName = "liability_item_id"
    End If
   
    ' Verify item exists
    Set rsCheck = db.OpenRecordset( _
        "SELECT " & fieldName & " FROM " & tableName & " WHERE " & fieldName & " = " & itemID)
   
    If rsCheck.EOF Then
        MsgBox IIf(isExpense, "Expense", "Liability") & " item not found!", vbCritical, "Error"
        rsCheck.Close
        Set db = Nothing
        Exit Function
    End If
    rsCheck.Close
    Set rsCheck = Nothing
   
    If prepaymentAmount <= 0 Then
        MsgBox "Prepayment amount must be greater than zero!", vbCritical, "Error"
        Set db = Nothing
        Exit Function
    End If
   
    ' Set default payment method if not provided
    If paymentMethodId <= 0 Then
        paymentMethodId = getDefaultPaymentMethodID()
    End If
   
    ' Determine prepayment date
    If IsNull(prepaymentDate) Then
        prepaymentDate = Now()
    End If
   
    ' Create payment record
    Set rsPayment = db.OpenRecordset("tbl_payments", dbOpenDynaset)
    rsPayment.AddNew
    
    ' Set the appropriate item ID field
    If isExpense Then
        rsPayment!expense_item_id = itemID
    Else
        rsPayment!liability_item_id = itemID
    End If
    
    rsPayment!payment_type = "Prepayment"
    rsPayment!payment_date = prepaymentDate
    rsPayment!amount_paid = prepaymentAmount
    rsPayment!payment_method = paymentMethodId
    If Len(description) > 0 Then rsPayment!description = description
    rsPayment.upDate
   
    rsPayment.Bookmark = rsPayment.LastModified
    newPaymentID = rsPayment!payment_id
    rsPayment.Close
    Set rsPayment = Nothing
   
    ' Create prepayment record
    Set rsPrepayment = db.OpenRecordset("tbl_prepayments", dbOpenDynaset)
    rsPrepayment.AddNew
    rsPrepayment!payment_id = newPaymentID
    
    ' Set the appropriate item ID field
    If isExpense Then
        rsPrepayment!expense_item_id = itemID
    Else
        rsPrepayment!liability_item_id = itemID
    End If
    
    rsPrepayment!total_prepaid = prepaymentAmount
    rsPrepayment!amount_utilized = 0
    'rsPrepayment!amount_remaining = prepaymentAmount 'THIS IS CALCULATED FIELD
    rsPrepayment!prepayment_date = prepaymentDate
    rsPrepayment!status = "Active"
    rsPrepayment.upDate
   
    rsPrepayment.Bookmark = rsPrepayment.LastModified
    newPrepaymentID = rsPrepayment!prepayment_id
    rsPrepayment.Close
    Set rsPrepayment = Nothing
    Set db = Nothing
   
    RecordPrepayment = newPrepaymentID
   
    MsgBox "Prepayment of " & Format(prepaymentAmount, "Currency") & " recorded!" & vbCrLf & _
           "Prepayment ID: " & newPrepaymentID, vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error recording prepayment: " & Err.description, vbCritical, "Error"
    On Error Resume Next
    If Not rsCheck Is Nothing Then rsCheck.Close
    If Not rsPayment Is Nothing Then rsPayment.Close
    If Not rsPrepayment Is Nothing Then rsPrepayment.Close
    Set rsCheck = Nothing
    Set rsPayment = Nothing
    Set rsPrepayment = Nothing
    Set db = Nothing
    RecordPrepayment = 0
End Function
' ================================================================================
' SECTION 6: AUTO-ALLOCATION (SAFE & ENHANCED)
' ================================================================================
' Update prepayment status based on actual utilization
Public Sub UpdatePrepaymentStatuses()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsPrepayments As DAO.Recordset
    Dim updatedCount As Integer
    Dim newStatus As String
    Dim actualRemaining As Currency
    
    Set db = CurrentDb
    updatedCount = 0
    
    Debug.Print "=== UpdatePrepaymentStatuses START ==="
    
    ' Get all prepayments regardless of current status
    Set rsPrepayments = db.OpenRecordset( _
        "SELECT prepayment_id, total_prepaid, amount_utilized, amount_remaining, status " & _
        "FROM tbl_prepayments", dbOpenDynaset)
    
    Do While Not rsPrepayments.EOF
        ' Calculate actual remaining amount
        actualRemaining = rsPrepayments!total_prepaid - rsPrepayments!amount_utilized
        
        ' Determine correct status based on actual remaining amount
        ' Utilized only when remaining is 0 or less, otherwise Active
        If actualRemaining <= 0 Then
            newStatus = "Utilized"
        Else
            newStatus = "Active"
        End If
        
        ' Update only if status has changed
        If rsPrepayments!status <> newStatus Then
            Debug.Print "Prepayment ID " & rsPrepayments!prepayment_id & _
                        " | Total: " & Format(rsPrepayments!total_prepaid, "Currency") & _
                        " | Utilized: " & Format(rsPrepayments!amount_utilized, "Currency") & _
                        " | Actual Remaining: " & Format(actualRemaining, "Currency") & _
                        " | Old Status: " & rsPrepayments!status & _
                        " | New Status: " & newStatus
            
            rsPrepayments.Edit
            rsPrepayments!status = newStatus
            rsPrepayments.upDate
            
            updatedCount = updatedCount + 1
        End If
        
        rsPrepayments.MoveNext
    Loop
    
    rsPrepayments.Close
    Set rsPrepayments = Nothing
    Set db = Nothing
    
    Debug.Print "=== UpdatePrepaymentStatuses END - Updated: " & updatedCount & " ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR in UpdatePrepaymentStatuses: " & Err.description
    On Error Resume Next
    If Not rsPrepayments Is Nothing Then rsPrepayments.Close
    Set rsPrepayments = Nothing
    Set db = Nothing
End Sub

Public Sub AutoAllocatePrepayments(Optional beforeDate As Variant = Null)
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rsObligations As DAO.Recordset
    Dim rsPrepayments As DAO.Recordset
    Dim rsAllocation As DAO.Recordset
    Dim allocationAmount As Currency
    Dim allocationsCount As Integer
    Dim newBalance As Currency
    Dim sqlWhere As String
    Dim cutoffDate As Date
    Dim actualRemaining As Currency
    Dim newRemaining As Currency
   
    Set db = CurrentDb
    allocationsCount = 0
   
    ' Determine cutoff date
    If IsNull(beforeDate) Then
        cutoffDate = Date
        Debug.Print "=== AutoAllocatePrepayments START (All obligations up to today) ==="
    Else
        cutoffDate = DateValue(CDate(beforeDate))
        Debug.Print "=== AutoAllocatePrepayments START (Obligations before " & Format(cutoffDate, "mm/dd/yyyy") & ") ==="
    End If
   
    ' Update all payment obligation statuses first
    UpdateAllPaymentStatuses
    
    ' Update all prepayment statuses to ensure accuracy
    UpdatePrepaymentStatuses
   
    ' Build SQL with date filter - include Pending, Partial, and Overdue
    sqlWhere = "SELECT obligation_id, expense_item_id, liability_item_id, balance, due_date, amount_due, amount_paid, prepayment_applied, payment_status " & _
               "FROM tbl_payment_obligations " & _
               "WHERE payment_status IN ('Pending', 'Partial', 'Overdue') AND balance > 0 " & _
               "AND DateValue(due_date) <= #" & Format(cutoffDate, "mm/dd/yyyy") & "# " & _
               "ORDER BY due_date"
   
    Debug.Print "SQL: " & sqlWhere
   
    Set rsObligations = db.OpenRecordset(sqlWhere)
   
    ' Move to last record to get actual count
    If Not rsObligations.EOF Then
        rsObligations.MoveLast
        rsObligations.MoveFirst
        Debug.Print "Found " & rsObligations.RecordCount & " obligations to process"
    Else
        Debug.Print "Found 0 obligations to process"
    End If
   
    Do While Not rsObligations.EOF
        Debug.Print "Checking Obligation ID " & rsObligations!obligation_id & _
                    " | Due: " & Format(DateValue(rsObligations!due_date), "mm/dd/yyyy") & _
                    " | Status: " & rsObligations!payment_status & _
                    " | Current Balance: " & Format(rsObligations!balance, "Currency")
       
        ' Determine which item ID to use - obligation can have EITHER expense OR liability, never both
        Dim itemID As Long
        Dim itemField As String
        Dim isExpenseItem As Boolean
        
        If Not IsNull(rsObligations!expense_item_id) Then
            itemID = rsObligations!expense_item_id
            itemField = "expense_item_id"
            isExpenseItem = True
            Debug.Print "  -> Using expense_item_id: " & itemID
        ElseIf Not IsNull(rsObligations!liability_item_id) Then
            itemID = rsObligations!liability_item_id
            itemField = "liability_item_id"
            isExpenseItem = False
            Debug.Print "  -> Using liability_item_id: " & itemID
        Else
            Debug.Print "  -> Skipping: No item ID found"
            rsObligations.MoveNext
            GoTo NextObligation
        End If
        
        ' Find prepayments for this item with actual remaining balance
        ' Don't trust status field - use calculation
        Dim prepaymentSQL As String
        prepaymentSQL = "SELECT prepayment_id, amount_remaining, amount_utilized, total_prepaid, status " & _
                       "FROM tbl_prepayments " & _
                       "WHERE " & itemField & " = " & itemID & _
                       " AND (total_prepaid - amount_utilized) > 0 " & _
                       "ORDER BY prepayment_date"
        
        Debug.Print "  -> Prepayment SQL: " & prepaymentSQL
        
        Set rsPrepayments = db.OpenRecordset(prepaymentSQL, dbOpenDynaset)
      
        If rsPrepayments.EOF Then
            Debug.Print "  -> No prepayments with remaining balance found for this item"
        Else
            ' Calculate actual remaining amount - don't trust the field
            actualRemaining = rsPrepayments!total_prepaid - rsPrepayments!amount_utilized
            
            Debug.Print "  -> Found prepayment ID " & rsPrepayments!prepayment_id & _
                        " | Total: " & Format(rsPrepayments!total_prepaid, "Currency") & _
                        " | Utilized: " & Format(rsPrepayments!amount_utilized, "Currency") & _
                        " | Actual Remaining: " & Format(actualRemaining, "Currency") & _
                        " | Status: " & rsPrepayments!status
            
            ' Determine allocation amount
            allocationAmount = IIf(actualRemaining >= rsObligations!balance, _
                                  rsObligations!balance, actualRemaining)
          
            Debug.Print "  -> Allocating " & Format(allocationAmount, "Currency") & _
                        " from Prepayment ID " & rsPrepayments!prepayment_id
          
            ' Record the allocation
            Set rsAllocation = db.OpenRecordset("tbl_payment_allocations", dbOpenDynaset)
            rsAllocation.AddNew
            rsAllocation!prepayment_id = rsPrepayments!prepayment_id
            rsAllocation!obligation_id = rsObligations!obligation_id
            rsAllocation!amount_allocated = allocationAmount
            rsAllocation!allocation_date = Date
            rsAllocation.upDate
            rsAllocation.Close
            Set rsAllocation = Nothing
          
            ' Update prepayment utilization
            rsPrepayments.Edit
            rsPrepayments!amount_utilized = rsPrepayments!amount_utilized + allocationAmount
            
            ' Calculate new remaining after allocation
            newRemaining = rsPrepayments!total_prepaid - (rsPrepayments!amount_utilized + allocationAmount)
            
            ' Update status based on remaining
            If newRemaining <= 0 Then
                rsPrepayments!status = "Utilized"
            Else
                rsPrepayments!status = "Active"
            End If
            
            Debug.Print "  -> New utilized: " & Format(rsPrepayments!amount_utilized + allocationAmount, "Currency") & _
                        " | New remaining: " & Format(newRemaining, "Currency") & _
                        " | New status: " & rsPrepayments!status
            rsPrepayments.upDate

            ' Update obligation
            rsObligations.Edit
            rsObligations!prepayment_applied = rsObligations!prepayment_applied + allocationAmount
            
            newBalance = rsObligations!amount_due - rsObligations!amount_paid - rsObligations!prepayment_applied
            
            If newBalance <= 0 Then
                rsObligations!payment_status = "Prepaid"
            Else
                rsObligations!payment_status = "Partial"
            End If
            rsObligations.upDate
          
            allocationsCount = allocationsCount + 1
        End If
      
        If Not rsPrepayments.EOF Then rsPrepayments.Close
        Set rsPrepayments = Nothing
        
NextObligation:
        rsObligations.MoveNext
    Loop
  
    rsObligations.Close
    Set rsObligations = Nothing
    Set db = Nothing
  
    Debug.Print "=== AutoAllocatePrepayments END - Allocations: " & allocationsCount & " ==="
    
    ' Update all statuses again after allocations
    UpdatePrepaymentStatuses
    UpdateAllPaymentStatuses
    
    If IsNull(beforeDate) Then
        Debug.Print "Successfully allocated " & allocationsCount & " prepayment(s) for obligations up to today", _
               vbInformation, "Allocation Complete"
    Else
        Debug.Print "Successfully allocated " & allocationsCount & " prepayment(s) for obligations before " & _
               Format(cutoffDate, "mm/dd/yyyy"), vbInformation, "Allocation Complete"
    End If
  
    Exit Sub
  
ErrorHandler:
    MsgBox "Error allocating prepayments: " & Err.description, vbCritical, "Error"
    Debug.Print "ERROR in AutoAllocatePrepayments: " & Err.description
    On Error Resume Next
    If Not rsObligations Is Nothing Then rsObligations.Close
    If Not rsPrepayments Is Nothing Then rsPrepayments.Close
    If Not rsAllocation Is Nothing Then rsAllocation.Close
    Set rsObligations = Nothing
    Set rsPrepayments = Nothing
    Set rsAllocation = Nothing
    Set db = Nothing
End Sub


' ================================================================================
' SECTION 7: STATUS MANAGEMENT (SAFE & ENHANCED)
' ================================================================================
Public Sub UpdateAllPaymentStatuses()
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rsObligations As DAO.Recordset
    Dim newStatus As String
    Dim oldStatus As String
    Dim recordsUpdated As Integer
   
    Set db = CurrentDb
    recordsUpdated = 0
   
    Debug.Print "=== UpdateAllPaymentStatuses START ==="
   
    Set rsObligations = db.OpenRecordset("SELECT * FROM tbl_payment_obligations", dbOpenDynaset)
   
    Do While Not rsObligations.EOF
        oldStatus = Nz(rsObligations!payment_status, "")
       
        If rsObligations!balance <= 0 Then
            newStatus = IIf(rsObligations!prepayment_applied > 0 And rsObligations!amount_paid = 0, "Prepaid", "Paid")
        ElseIf rsObligations!due_date < Date And rsObligations!balance > 0 Then
            newStatus = "Overdue"
        ElseIf (rsObligations!amount_paid > 0 Or rsObligations!prepayment_applied > 0) And rsObligations!balance > 0 Then
            newStatus = "Partial"
        Else
            newStatus = "Pending"
        End If
       
        If newStatus <> oldStatus Then
            rsObligations.Edit
            rsObligations!payment_status = newStatus
            rsObligations.upDate
            recordsUpdated = recordsUpdated + 1
            Debug.Print "  Obligation " & rsObligations!obligation_id & ": " & oldStatus & " ? " & newStatus
        End If
       
        rsObligations.MoveNext
    Loop
   
    rsObligations.Close
    Set rsObligations = Nothing
    Set db = Nothing
   
    Debug.Print "=== UpdateAllPaymentStatuses END - Updated: " & recordsUpdated & " ==="
    'MsgBox "Updated " & recordsUpdated & " payment status(es)", vbInformation, "Update Complete"
   
    Exit Sub
   
ErrorHandler:
    MsgBox "Error updating statuses: " & Err.description, vbCritical, "Error"
    Debug.Print "ERROR in UpdateAllPaymentStatuses: " & Err.description
End Sub

' ================================================================================
' SECTION 8: OBLIGATION MANAGEMENT (SAFE)
' ================================================================================
Public Function CreateManualObligation(expenseItemID As Long, _
                                       dueDate As Date, _
                                       amountDue As Currency, _
                                       Optional description As String = "") As Long
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim newID As Long
   
    Set db = CurrentDb
    CreateManualObligation = 0
   
    Set rs = db.OpenRecordset("tbl_payment_obligations", dbOpenDynaset)
    rs.AddNew
    rs!expense_item_id = expenseItemID
    rs!obligation_type = "Expense"
    rs!obligation_date = Date
    rs!due_date = dueDate
    rs!amount_due = amountDue
    rs!amount_paid = 0
    rs!prepayment_applied = 0
    rs!payment_status = "Pending"
    If Len(description) > 0 Then rs!description = description
    rs.upDate
   
    rs.Bookmark = rs.LastModified
    newID = rs!obligation_id
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    CreateManualObligation = newID
    MsgBox "Obligation created successfully!" & vbCrLf & "Obligation ID: " & newID, _
           vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error creating obligation: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    CreateManualObligation = 0
End Function

Public Function UpdateObligation(obligationID As Long, _
                                 Optional dueDate As Variant = Null, _
                                 Optional amountDue As Variant = Null, _
                                 Optional description As Variant = Null) As Boolean
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
   
    Set db = CurrentDb
    UpdateObligation = False
   
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tbl_payment_obligations WHERE obligation_id = " & obligationID, _
        dbOpenDynaset)
   
    If rs.EOF Then
        MsgBox "Obligation not found!", vbCritical, "Error"
        rs.Close
        Set db = Nothing
        Exit Function
    End If
   
    rs.Edit
    If IsDate(dueDate) Then rs!due_date = CDate(dueDate)
    If Not IsNull(amountDue) Then rs!amount_due = CCur(amountDue)
    If Not IsNull(description) Then rs!description = description
    rs.upDate
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    UpdateObligation = True
    MsgBox "Obligation updated successfully!", vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error updating obligation: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    UpdateObligation = False
End Function

Public Function DeleteObligation(obligationID As Long) As Boolean
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
   
    Set db = CurrentDb
    DeleteObligation = False
   
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS cnt FROM tbl_payments WHERE obligation_id = " & obligationID)
   
    If Not rs.EOF Then
        If rs!cnt > 0 Then
            MsgBox "Cannot delete: This obligation has " & rs!cnt & " payments!", _
                   vbCritical, "Delete Failed"
            rs.Close
            Set db = Nothing
            Exit Function
        End If
    End If
    rs.Close
   
    If MsgBox("Are you sure you want to delete this obligation?", _
              vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Set db = Nothing
        Exit Function
    End If
   
    db.Execute "DELETE FROM tbl_payment_allocations WHERE obligation_id = " & obligationID
   
    db.Execute "DELETE FROM tbl_payment_obligations WHERE obligation_id = " & obligationID
   
    Set db = Nothing
    DeleteObligation = True
    MsgBox "Obligation deleted successfully!", vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error deleting obligation: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    DeleteObligation = False
End Function

' ================================================================================
' SECTION 9: PAYMENT RECORDING (SAFE)
' ================================================================================
Public Function RecordExpensePayment(obligationID As Long, _
                                     Optional paymentAmount As Currency = 0, _
                                     Optional paymentDate As Variant = Null, _
                                     Optional paymentMethodId As Long = 0, _
                                     Optional referenceNum As String = "", _
                                     Optional useDueDate As Boolean = False, _
                                     Optional payWholeOnTime As Boolean = False _
                                     ) As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsObligation As DAO.Recordset
    Dim rsPayment As DAO.Recordset
    Dim newPaymentID As Long
    Dim tempBalance As Currency
    Dim expenseItemID As Variant
    Dim liabilityItemID As Variant
   
    Set db = CurrentDb
    RecordExpensePayment = 0
   
    ' Open obligation record to get expense_item_id and liability_item_id
    Set rsObligation = db.OpenRecordset( _
        "SELECT * FROM tbl_payment_obligations WHERE obligation_id = " & obligationID, _
        dbOpenDynaset)
   
    If rsObligation.EOF Then
        MsgBox "Obligation not found!", vbCritical, "Error"
        rsObligation.Close
        Set db = Nothing
        Exit Function
    End If
    
    ' Get the item IDs from the obligation record
    Dim amount_due As Currency
    expenseItemID = rsObligation!expense_item_id
    liabilityItemID = rsObligation!liability_item_id
    amount_due = rsObligation!amount_due
    
    If paymentMethodId <= 0 Then
        paymentMethodId = getDefaultPaymentMethodID()
    End If
   
   'Pay as it was Expected?
    If payWholeOnTime Then
        useDueDate = True
        paymentAmount = amount_due
    End If
    
    ' Determine payment date
    If useDueDate Then
        paymentDate = rsObligation!due_date
    ElseIf IsNull(paymentDate) Then
        paymentDate = Now()
    End If
    

   
    If paymentAmount <= 0 Then
        MsgBox "Payment amount must be greater than zero!", vbCritical, "Error"
        rsObligation.Close
        Set db = Nothing
        Exit Function
    End If
   
    ' Create payment record
    Set rsPayment = db.OpenRecordset("tbl_payments", dbOpenDynaset)
    rsPayment.AddNew
    rsPayment!obligation_id = obligationID
    
    ' Set expense_item_id if it exists in the obligation
    If Not IsNull(expenseItemID) Then
        rsPayment!expense_item_id = expenseItemID
    End If
    
    ' Set liability_item_id if it exists in the obligation
    If Not IsNull(liabilityItemID) Then
        rsPayment!liability_item_id = liabilityItemID
    End If
    
    rsPayment!payment_type = "Regular"
    rsPayment!payment_date = paymentDate
    rsPayment!amount_paid = paymentAmount
    rsPayment!payment_method = paymentMethodId
    If Len(referenceNum) > 0 Then rsPayment!reference_number = referenceNum
    rsPayment.upDate
   
    rsPayment.Bookmark = rsPayment.LastModified
    newPaymentID = rsPayment!payment_id
    rsPayment.Close
    Set rsPayment = Nothing
   
    ' Update obligation record
    rsObligation.Edit
    rsObligation!amount_paid = rsObligation!amount_paid + paymentAmount
    
    tempBalance = rsObligation!amount_due - rsObligation!amount_paid - rsObligation!prepayment_applied
    
    If tempBalance <= 0 Then
        rsObligation!payment_status = "Paid"
    Else
        rsObligation!payment_status = "Partial"
    End If
    rsObligation.upDate
   
    rsObligation.Close
    Set rsObligation = Nothing
    Set db = Nothing
   
    RecordExpensePayment = newPaymentID
    MsgBox "Payment recorded successfully!" & vbCrLf & "Payment ID: " & newPaymentID, _
           vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error recording payment: " & Err.description, vbCritical, "Error"
    On Error Resume Next
    If Not rsObligation Is Nothing Then rsObligation.Close
    If Not rsPayment Is Nothing Then rsPayment.Close
    Set rsObligation = Nothing
    Set rsPayment = Nothing
    Set db = Nothing
    RecordExpensePayment = 0
End Function


Public Function RecordLoanPayment(liabilityItemID As Long, _
                                  principalAmount As Currency, _
                                  interestAmount As Currency, _
                                  paymentMethodId As Long, _
                                  Optional paymentDate As Variant = "", _
                                  Optional referenceNum As String = "") As Long
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rsLiability As DAO.Recordset
    Dim rsPayment As DAO.Recordset
    Dim rsDetail As DAO.Recordset
    Dim currentBalance As Currency
    Dim totalPayment As Currency
    Dim newBalance As Currency
    Dim newPaymentID As Long
   
    Set db = CurrentDb
    RecordLoanPayment = 0
    If IsDate(paymentDate) Then
        paymentDate = CDate(paymentDate)
    Else
        paymentDate = Now()
    End If
   
    Set rsLiability = db.OpenRecordset( _
        "SELECT * FROM tbl_liability_items WHERE liability_item_id = " & liabilityItemID, _
        dbOpenDynaset)
   
    If rsLiability.EOF Then
        MsgBox "Liability not found!", vbCritical, "Error"
        rsLiability.Close
        Set db = Nothing
        Exit Function
    End If
   
    currentBalance = rsLiability!current_balance
    totalPayment = principalAmount + interestAmount
   
    If totalPayment <= 0 Then
        MsgBox "Total payment must be greater than zero!", vbCritical, "Error"
        rsLiability.Close
        Set db = Nothing
        Exit Function
    End If
   
    If principalAmount > currentBalance Then
        MsgBox "Principal exceeds current balance!", vbCritical, "Error"
        rsLiability.Close
        Set db = Nothing
        Exit Function
    End If
   
    newBalance = currentBalance - principalAmount
   
    Set rsPayment = db.OpenRecordset("tbl_payments", dbOpenDynaset)
    rsPayment.AddNew
    rsPayment!liability_item_id = liabilityItemID
    rsPayment!payment_type = "Regular"
    rsPayment!payment_date = paymentDate
    rsPayment!amount_paid = totalPayment
    rsPayment!payment_method = paymentMethodId
    If Len(referenceNum) > 0 Then rsPayment!reference_number = referenceNum
    rsPayment.upDate
   
    rsPayment.Bookmark = rsPayment.LastModified
    newPaymentID = rsPayment!payment_id
    rsPayment.Close
    Set rsPayment = Nothing
   
    Set rsDetail = db.OpenRecordset("tbl_liability_payment_details", dbOpenDynaset)
    rsDetail.AddNew
    rsDetail!payment_id = newPaymentID
    rsDetail!liability_item_id = liabilityItemID
    rsDetail!principal_amount = principalAmount
    rsDetail!interest_amount = interestAmount
    rsDetail!balance_after_payment = newBalance
    rsDetail!payment_date = paymentDate
    rsDetail.upDate
    rsDetail.Close
    Set rsDetail = Nothing
   
    rsLiability.Edit
    rsLiability!current_balance = newBalance
    rsLiability.upDate
    rsLiability.Close
    Set rsLiability = Nothing
    Set db = Nothing
   
    RecordLoanPayment = newPaymentID
    MsgBox "Loan payment recorded!" & vbCrLf & _
           "Payment ID: " & newPaymentID & vbCrLf & _
           "New Balance: " & Format(newBalance, "Currency"), _
           vbInformation, "Success"
   
    Exit Function
   
ErrorHandler:
    MsgBox "Error recording loan payment: " & Err.description, vbCritical, "Error"
    On Error Resume Next
    If Not rsLiability Is Nothing Then rsLiability.Close
    If Not rsPayment Is Nothing Then rsPayment.Close
    If Not rsDetail Is Nothing Then rsDetail.Close
    Set rsLiability = Nothing
    Set rsPayment = Nothing
    Set rsDetail = Nothing
    Set db = Nothing
    RecordLoanPayment = 0
End Function

Public Function DeletePayment(paymentID As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim amount As Currency
    Dim liabilityItemID As Long
    Dim currentBalance As Currency
   
    Dim db As DAO.Database
   
    Set db = CurrentDb
    DeletePayment = False
   
    If MsgBox("Are you sure you want to delete this payment?" & vbCrLf & _
              "This will also delete payment details and affect balances!", _
              vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Set db = Nothing
        Exit Function
    End If
   
    db.Execute "DELETE FROM tbl_liability_payment_details WHERE payment_id = " & paymentID
   
    amount = Nz(DLookup("amount_paid", "tbl_payments", "payment_id=" & paymentID), 0)
    liabilityItemID = Nz(DLookup("liability_item_id", "tbl_payments", "payment_id=" & paymentID), 0)
    currentBalance = Nz(DLookup("current_balance", "tbl_liability_items", "liability_item_id=" & liabilityItemID), 0)
    db.Execute "DELETE FROM tbl_payments WHERE payment_id = " & paymentID
   
    If liabilityItemID > 0 Then
        db.Execute "UPDATE tbl_liability_items SET current_balance = " & (currentBalance + amount) & _
                   " WHERE liability_item_id = " & liabilityItemID
    End If
   
    Set db = Nothing
    DeletePayment = True
    MsgBox "Payment deleted! Please run maintenance.", vbInformation, "Success"
    RunDailyMaintenance
    Exit Function
   
ErrorHandler:
    MsgBox "Error deleting payment: " & Err.description, vbCritical, "Error"
    Set db = Nothing
    DeletePayment = False
End Function

' ================================================================================
' SECTION 10: REPORTING & UTILITY FUNCTIONS
' ================================================================================
Public Function GetTotalOutstandingBalance() As Currency
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim total As Currency
   
    Set db = CurrentDb
    total = 0
   
    Set rs = db.OpenRecordset( _
        "SELECT SUM(balance) AS total_balance FROM tbl_payment_obligations " & _
        "WHERE payment_status IN ('Pending', 'Partial', 'Overdue')")
   
    If Not rs.EOF Then
        total = Nz(rs!total_balance, 0)
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetTotalOutstandingBalance = total
End Function

Public Function GetOverdueCount() As Long
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim count As Long
   
    Set db = CurrentDb
    count = 0
   
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS overdue_count FROM tbl_payment_obligations " & _
        "WHERE payment_status = 'Overdue'")
   
    If Not rs.EOF Then
        count = Nz(rs!overdue_count, 0)
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetOverdueCount = count
End Function

Public Function GetAvailablePrepayment(expenseItemID As Long) As Currency
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim available As Currency
   
    Set db = CurrentDb
    available = 0
   
    Set rs = db.OpenRecordset( _
        "SELECT SUM(amount_remaining) AS total_available FROM tbl_prepayments " & _
        "WHERE expense_item_id = " & expenseItemID & " AND status = 'Active'")
   
    If Not rs.EOF Then
        available = Nz(rs!total_available, 0)
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetAvailablePrepayment = available
End Function

Public Function GetTotalLiabilityBalance() As Currency
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim total As Currency
   
    Set db = CurrentDb
    total = 0
   
    Set rs = db.OpenRecordset( _
        "SELECT SUM(current_balance) AS total_balance FROM tbl_liability_items " & _
        "WHERE is_active = True")
   
    If Not rs.EOF Then
        total = Nz(rs!total_balance, 0)
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetTotalLiabilityBalance = total
End Function

Public Function GetExpenseItemName(expenseItemID As Long) As String
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim itemName As String
   
    Set db = CurrentDb
    itemName = "Unknown"
   
    Set rs = db.OpenRecordset( _
        "SELECT item_name FROM tbl_expense_items WHERE expense_item_id = " & expenseItemID)
   
    If Not rs.EOF Then
        itemName = Nz(rs!item_name, "Unknown")
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetExpenseItemName = itemName
End Function

Public Function GetLiabilityItemName(liabilityItemID As Long) As String
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim itemName As String
   
    Set db = CurrentDb
    itemName = "Unknown"
   
    Set rs = db.OpenRecordset( _
        "SELECT liability_name FROM tbl_liability_items WHERE liability_item_id = " & liabilityItemID)
   
    If Not rs.EOF Then
        itemName = Nz(rs!liability_name, "Unknown")
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetLiabilityItemName = itemName
End Function

Public Function GetPaymentMethodName(paymentMethodId As Long) As String
    On Error Resume Next
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim methodName As String
   
    Set db = CurrentDb
    methodName = "Unknown"
   
    Set rs = db.OpenRecordset( _
        "SELECT payment_method FROM tbl_payment_methods WHERE payment_method_id = " & paymentMethodId)
   
    If Not rs.EOF Then
        methodName = Nz(rs!payment_method, "Unknown")
    End If
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    GetPaymentMethodName = methodName
End Function

Public Sub GenerateFinancialSummary()
    On Error GoTo ErrorHandler
   
    Dim totalOutstanding As Currency
    Dim totalLiabilities As Currency
    Dim overdueCount As Long
    Dim reportText As String
   
    totalOutstanding = GetTotalOutstandingBalance()
    totalLiabilities = GetTotalLiabilityBalance()
    overdueCount = GetOverdueCount()
   
    reportText = "FINANCIAL SUMMARY" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    reportText = reportText & "Outstanding Obligations: " & Format(totalOutstanding, "Currency") & vbCrLf
    reportText = reportText & "Total Liabilities: " & Format(totalLiabilities, "Currency") & vbCrLf
    reportText = reportText & "Overdue Obligations: " & overdueCount & vbCrLf & vbCrLf
    reportText = reportText & "Total Exposure: " & Format(totalOutstanding + totalLiabilities, "Currency")
   
    MsgBox reportText, vbInformation, "Financial Summary"
   
    Exit Sub
   
ErrorHandler:
    MsgBox "Error generating summary: " & Err.description, vbCritical, "Error"
End Sub

' ================================================================================
' SECTION 11: MASTER MAINTENANCE FUNCTIONS
' ================================================================================
Public Sub RunDailyMaintenance()
    On Error Resume Next
   
    UpdateAllPaymentStatuses
    AutoAllocatePrepayments
   
    MsgBox "Daily maintenance completed!", vbInformation, "Maintenance Complete"
End Sub

Public Sub GenerateAllObligations(Optional targetDate As Variant = Null)
    Call GenerateObligations("Monthly", targetDate)
    Call GenerateObligations("Weekly", targetDate)
    Call AutoAllocatePrepayments
   
    'MsgBox "All obligations generated and prepayments allocated!", vbInformation, "Complete"
End Sub

Public Sub UpdateLiabilityBalance()
    On Error GoTo ErrorHandler
   
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim updatedCount As Integer
    Dim principal As Currency
    Dim interest As Currency
   
    Set db = CurrentDb
    updatedCount = 0
   
    Set rs = db.OpenRecordset("SELECT * FROM tbl_liability_items", dbOpenDynaset)
   
    Do While Not rs.EOF
        principal = Nz(DSum("principal_amount", "tbl_liability_payment_details", "liability_item_id=" & rs!liability_item_id), 0)
        interest = Nz(DSum("interest_amount", "tbl_liability_payment_details", "liability_item_id=" & rs!liability_item_id), 0)
        If rs.Updatable Then
            rs.Edit
            rs!current_balance = rs!original_amount - principal - interest
            rs.upDate
        End If
        updatedCount = updatedCount + 1
        rs.MoveNext
    Loop
   
    rs.Close
    Set rs = Nothing
    Set db = Nothing
   
    UpdateAllPaymentStatuses
   
    MsgBox "Recalculated " & updatedCount & " liability balances", vbInformation, "Recalculation Complete"
   
    Exit Sub
   
ErrorHandler:
    MsgBox "Error recalculating balances: " & Err.description, vbCritical, "Error"
    If Not rs Is Nothing Then If Not rs.EOF Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub




' Generate obligations for a range of dates based on recurrence type
Public Sub GenerateHistoricalObligations(startDate As Date, _
                                         EndDate As Date, _
                                         recurrenceType As String)
    On Error GoTo ErrorHandler
    
    Dim currentDate As Date
    Dim periodsGenerated As Integer
    Dim intervalType As String
    Dim intervalAmount As Integer
    
    currentDate = startDate
    periodsGenerated = 0
    
    ' Determine interval based on recurrence type
    Select Case LCase(recurrenceType)
        Case "daily"
            intervalType = "d"
            intervalAmount = 1
        Case "weekly"
            intervalType = "ww"
            intervalAmount = 1
        Case "monthly"
            intervalType = "m"
            intervalAmount = 1
        Case "annually", "yearly", "annual"
            intervalType = "yyyy"
            intervalAmount = 1
        Case Else
            MsgBox "Invalid recurrence type: " & recurrenceType & vbCrLf & _
                   "Valid types: Daily, Weekly, Monthly, Annually/Yearly", _
                   vbCritical, "Error"
            Exit Sub
    End Select
    
    ' Generate obligations for each period
    Do While currentDate <= EndDate
        Call GenerateAllObligations(currentDate)
        periodsGenerated = periodsGenerated + 1
        
        ' Move to next period
        currentDate = DateAdd(intervalType, intervalAmount, currentDate)
    Loop
    
    MsgBox "Generated " & periodsGenerated & " " & recurrenceType & _
           " obligations from " & Format(startDate, "mm/dd/yyyy") & _
           " to " & Format(EndDate, "mm/dd/yyyy") & "!", _
           vbInformation, "Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating historical obligations: " & Err.description, _
           vbCritical, "Error"
End Sub


