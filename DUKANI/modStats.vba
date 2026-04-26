Option Compare Database
Option Explicit

' ============================================================================
' MODULE: modStats
' PURPOSE: Basic aggregation layer for quantities and monetary amounts
' DESCRIPTION: Data access functions for retrieving totals from database
' 
' STOCK CALCULATION FORMULA:
'   Current Stock = Purchases + Return Inwards - Sales - Credit Sales 
'                   - Return Outwards - Office Use - Drawings
'
' TRANSACTION INTEGRATION:
'   - Purchases: Increase inventory, increase purchase cost
'   - Sales (Cash): Decrease inventory, increase sales revenue
'   - Credit Sales (Debts): Decrease inventory, increase sales revenue + receivables
'   - Return Inwards: Increase inventory, decrease sales revenue
'   - Return Outwards: Decrease inventory, decrease purchase cost
'   - Office Use: Decrease inventory, expensed as operating cost
'   - Drawings: Decrease inventory, reduces owner's equity
' ============================================================================

Public Function getSales( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal brand_id As Long = 0, _
    Optional ByVal category_id As Long = 0, _
    Optional ByVal type_id As Long = 0, _
    Optional ByVal month_id As Long = 0, _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null, _
    Optional ByVal returnRecordset As Boolean = False) As Variant
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim total As Currency
    
    Debug.Print "  [getSales] product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "  [getSales] month_id: " & month_id & ", start_date: " & start_date & ", end_date: " & end_date
    
    Set db = CurrentDb
    sql = "SELECT amount, sale_date FROM qry_sales WHERE 1=1"
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        sql = sql & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        sql = sql & " AND product_spec_id = " & product_spec_id
    End If
    If brand_id > 0 Then sql = sql & " AND brand_id = " & brand_id
    If category_id > 0 Then sql = sql & " AND category_id = " & category_id
    If type_id > 0 Then sql = sql & " AND type_id = " & type_id
    If month_id > 0 Then sql = sql & " AND month_id = " & month_id
    If Not IsNull(start_date) And IsDate(start_date) Then sql = sql & " AND sale_date >= #" & Format(start_date, "mm/dd/yyyy") & "#"
    If Not IsNull(end_date) And IsDate(end_date) Then sql = sql & " AND sale_date <= #" & Format(end_date, "mm/dd/yyyy") & "#"
    sql = sql & " ORDER BY sale_date DESC"
    
    Debug.Print "  [getSales] SQL: " & sql
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If returnRecordset Then
        Set getSales = rs
    Else
        total = 0
        Do While Not rs.EOF
            total = total + Nz(rs!amount, 0)
            rs.MoveNext
        Loop
        rs.Close
        Debug.Print "  [getSales] Total: " & Format(total, "#,##0.00")
        getSales = total
    End If
End Function

Public Function getReturnInwards( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal brand_id As Long = 0, _
    Optional ByVal category_id As Long = 0, _
    Optional ByVal type_id As Long = 0, _
    Optional ByVal month_id As Long = 0, _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null, _
    Optional ByVal returnRecordset As Boolean = False) As Variant
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim total As Currency
    
    Debug.Print "  [getReturnInwards] product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "  [getReturnInwards] month_id: " & month_id & ", start_date: " & start_date & ", end_date: " & end_date
    
    Set db = CurrentDb
    sql = "SELECT amount, sale_date FROM qry_return_inwards WHERE 1=1"
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        sql = sql & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        sql = sql & " AND product_spec_id = " & product_spec_id
    End If
    If brand_id > 0 Then sql = sql & " AND brand_id = " & brand_id
    If category_id > 0 Then sql = sql & " AND category_id = " & category_id
    If type_id > 0 Then sql = sql & " AND type_id = " & type_id
    If month_id > 0 Then sql = sql & " AND month_id = " & month_id
    If Not IsNull(start_date) And IsDate(start_date) Then sql = sql & " AND sale_date >= #" & Format(start_date, "mm/dd/yyyy") & "#"
    If Not IsNull(end_date) And IsDate(end_date) Then sql = sql & " AND sale_date <= #" & Format(end_date, "mm/dd/yyyy") & "#"
    sql = sql & " ORDER BY sale_date DESC"
    
    Debug.Print "  [getReturnInwards] SQL: " & sql
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If returnRecordset Then
        Set getReturnInwards = rs
    Else
        total = 0
        Do While Not rs.EOF
            total = total + Nz(rs!amount, 0)
            rs.MoveNext
        Loop
        rs.Close
        Debug.Print "  [getReturnInwards] Total: " & Format(total, "#,##0.00")
        getReturnInwards = total
    End If
End Function

Public Function getNetSales( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal brand_id As Long = 0, _
    Optional ByVal category_id As Long = 0, _
    Optional ByVal type_id As Long = 0, _
    Optional ByVal month_id As Long = 0, _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null) As Currency
    
    Dim salesTotal As Currency
    Dim returnInwardsTotal As Currency
    Dim result As Currency
    
    Debug.Print "=== getNetSales DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", start_date: " & start_date & ", end_date: " & end_date
    
    salesTotal = getSales(product_spec_id, product_id, brand_id, category_id, type_id, month_id, start_date, end_date)
    Debug.Print "Sales Total: " & Format(salesTotal, "#,##0.00")
    
    returnInwardsTotal = getReturnInwards(product_spec_id, product_id, brand_id, category_id, type_id, month_id, start_date, end_date)
    Debug.Print "Return Inwards Total: " & Format(returnInwardsTotal, "#,##0.00")
    
    result = salesTotal - returnInwardsTotal
    Debug.Print "Net Sales (Sales - Return Inwards): " & Format(result, "#,##0.00")
    Debug.Print "=== END getNetSales DEBUG ==="
    
    getNetSales = result
End Function

Public Function getPurchases( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal brand_id As Long = 0, _
    Optional ByVal category_id As Long = 0, _
    Optional ByVal type_id As Long = 0, _
    Optional ByVal month_id As Long = 0, _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null, _
    Optional ByVal returnRecordset As Boolean = False) As Variant
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim total As Currency
    
    Debug.Print "  [getPurchases] product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "  [getPurchases] month_id: " & month_id & ", start_date: " & start_date & ", end_date: " & end_date
    
    Set db = CurrentDb
    sql = "SELECT amount, purchase_date FROM qry_purchases WHERE 1=1"
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        sql = sql & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        sql = sql & " AND product_spec_id = " & product_spec_id
    End If
    If brand_id > 0 Then sql = sql & " AND brand_id = " & brand_id
    If category_id > 0 Then sql = sql & " AND category_id = " & category_id
    If type_id > 0 Then sql = sql & " AND type_id = " & type_id
    If month_id > 0 Then sql = sql & " AND month_id = " & month_id
    If Not IsNull(start_date) And IsDate(start_date) Then sql = sql & " AND purchase_date >= #" & Format(start_date, "mm/dd/yyyy") & "#"
    If Not IsNull(end_date) And IsDate(end_date) Then sql = sql & " AND purchase_date <= #" & Format(end_date, "mm/dd/yyyy") & "#"
    sql = sql & " ORDER BY purchase_date DESC"
    
    Debug.Print "  [getPurchases] SQL: " & sql
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If returnRecordset Then
        Set getPurchases = rs
    Else
        total = 0
        Do While Not rs.EOF
            total = total + Nz(rs!amount, 0)
            rs.MoveNext
        Loop
        rs.Close
        Debug.Print "  [getPurchases] Total: " & Format(total, "#,##0.00")
        getPurchases = total
    End If
End Function

Public Function getReturnOutwards( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal brand_id As Long = 0, _
    Optional ByVal category_id As Long = 0, _
    Optional ByVal type_id As Long = 0, _
    Optional ByVal month_id As Long = 0, _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null, _
    Optional ByVal returnRecordset As Boolean = False) As Variant
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim total As Currency
    
    Debug.Print "  [getReturnOutwards] product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "  [getReturnOutwards] month_id: " & month_id & ", start_date: " & start_date & ", end_date: " & end_date
    
    Set db = CurrentDb
    sql = "SELECT amount, purchase_date FROM qry_return_outwards WHERE 1=1"
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        sql = sql & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        sql = sql & " AND product_spec_id = " & product_spec_id
    End If
    If brand_id > 0 Then sql = sql & " AND brand_id = " & brand_id
    If category_id > 0 Then sql = sql & " AND category_id = " & category_id
    If type_id > 0 Then sql = sql & " AND type_id = " & type_id
    If month_id > 0 Then sql = sql & " AND month_id = " & month_id
    If Not IsNull(start_date) And IsDate(start_date) Then sql = sql & " AND purchase_date >= #" & Format(start_date, "mm/dd/yyyy") & "#"
    If Not IsNull(end_date) And IsDate(end_date) Then sql = sql & " AND purchase_date <= #" & Format(end_date, "mm/dd/yyyy") & "#"
    sql = sql & " ORDER BY purchase_date DESC"
    
    Debug.Print "  [getReturnOutwards] SQL: " & sql
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If returnRecordset Then
        Set getReturnOutwards = rs
    Else
        total = 0
        Do While Not rs.EOF
            total = total + Nz(rs!amount, 0)
            rs.MoveNext
        Loop
        rs.Close
        Debug.Print "  [getReturnOutwards] Total: " & Format(total, "#,##0.00")
        getReturnOutwards = total
    End If
End Function

Public Function getNetPurchases( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal brand_id As Long = 0, _
    Optional ByVal category_id As Long = 0, _
    Optional ByVal type_id As Long = 0, _
    Optional ByVal month_id As Long = 0, _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null) As Currency
    
    Dim purchasesTotal As Currency
    Dim returnOutwardsTotal As Currency
    Dim result As Currency
    
    Debug.Print "=== getNetPurchases DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", start_date: " & start_date & ", end_date: " & end_date
    
    purchasesTotal = getPurchases(product_spec_id, product_id, brand_id, category_id, type_id, month_id, start_date, end_date)
    Debug.Print "Purchases Total: " & Format(purchasesTotal, "#,##0.00")
    
    returnOutwardsTotal = getReturnOutwards(product_spec_id, product_id, brand_id, category_id, type_id, month_id, start_date, end_date)
    Debug.Print "Return Outwards Total: " & Format(returnOutwardsTotal, "#,##0.00")
    
    result = purchasesTotal - returnOutwardsTotal
    Debug.Print "Net Purchases (Purchases - Return Outwards): " & Format(result, "#,##0.00")
    Debug.Print "=== END getNetPurchases DEBUG ==="
    
    getNetPurchases = result
End Function

Public Function getExpenses( _
    Optional ByVal expense_type_id As Long = 0, _
    Optional ByVal month_id As Long = 0, _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null) As Currency
    
    Dim sql As String
    sql = "SELECT Sum(amount) AS total FROM tbl_expenses WHERE 1=1"
    If expense_type_id > 0 Then sql = sql & " AND expense_type_id = " & expense_type_id
    If month_id > 0 Then sql = sql & " AND Month(expense_date) = " & month_id & " AND Year(expense_date) = Year(Date())"
    If Not IsNull(start_date) And IsDate(start_date) Then sql = sql & " AND expense_date >= #" & Format(start_date, "mm/dd/yyyy") & "#"
    If Not IsNull(end_date) And IsDate(end_date) Then sql = sql & " AND expense_date <= #" & Format(end_date, "mm/dd/yyyy") & "#"
    
    getExpenses = Nz(CurrentDb.OpenRecordset(sql)(0), 0)
End Function

Public Function getReceivablesOutstanding( _
    Optional ByVal start_date As Variant = Null, _
    Optional ByVal end_date As Variant = Null) As Currency
    
    Dim sql As String
    sql = "SELECT Sum(amount - Nz(paid_amount,0)) AS outstanding FROM qry_debts_outstanding WHERE 1=1"
    If Not IsNull(start_date) And IsDate(start_date) Then sql = sql & " AND sale_date >= #" & Format(start_date, "mm/dd/yyyy") & "#"
    If Not IsNull(end_date) And IsDate(end_date) Then sql = sql & " AND sale_date <= #" & Format(end_date, "mm/dd/yyyy") & "#"
    
    getReceivablesOutstanding = Nz(CurrentDb.OpenRecordset(sql)(0), 0)
End Function

' Quantity-based stock functions with all adjustments
Public Function getPurchasedQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim cond As String: cond = "1=1"
    Dim result As Long
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        cond = cond & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        cond = cond & " AND product_spec_id = " & product_spec_id
    End If
    If IsDate(up_to_date) Then cond = cond & " AND purchase_date <= #" & Format(up_to_date, "mm/dd/yyyy") & "#"
    
    Debug.Print "  [getPurchasedQuantity] Condition: " & cond
    result = Nz(DSum("quantity", "qry_purchases", cond), 0)
    Debug.Print "  [getPurchasedQuantity] Result: " & result
    
    getPurchasedQuantity = result
End Function

Public Function getSoldQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim cond As String: cond = "1=1"
    Dim result As Long
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        cond = cond & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        cond = cond & " AND product_spec_id = " & product_spec_id
    End If
    If IsDate(up_to_date) Then cond = cond & " AND sale_date <= #" & Format(up_to_date, "mm/dd/yyyy") & "#"
    
    Debug.Print "  [getSoldQuantity] Condition: " & cond
    result = Nz(DSum("quantity", "qry_sales", cond), 0)
    Debug.Print "  [getSoldQuantity] Result: " & result
    
    getSoldQuantity = result
End Function

Public Function getOfficeUseQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim cond As String: cond = "1=1"
    Dim result As Long
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        cond = cond & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        cond = cond & " AND product_spec_id = " & product_spec_id
    End If
    If IsDate(up_to_date) Then cond = cond & " AND sale_date <= #" & Format(up_to_date, "mm/dd/yyyy") & "#"
    
    Debug.Print "  [getOfficeUseQuantity] Condition: " & cond
    result = Nz(DSum("quantity", "qry_sales_office_use", cond), 0)
    Debug.Print "  [getOfficeUseQuantity] Result: " & result
    
    getOfficeUseQuantity = result
End Function

Public Function getDrawingsQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim cond As String: cond = "1=1"
    Dim result As Long
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        cond = cond & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        cond = cond & " AND product_spec_id = " & product_spec_id
    End If
    If IsDate(up_to_date) Then cond = cond & " AND sale_date <= #" & Format(up_to_date, "mm/dd/yyyy") & "#"
    
    Debug.Print "  [getDrawingsQuantity] Condition: " & cond
    result = Nz(DSum("quantity", "qry_drawings", cond), 0)
    Debug.Print "  [getDrawingsQuantity] Result: " & result
    
    getDrawingsQuantity = result
End Function

Public Function getReturnInwardsQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim cond As String: cond = "1=1"
    Dim result As Long
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        cond = cond & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        cond = cond & " AND product_spec_id = " & product_spec_id
    End If
    If IsDate(up_to_date) Then cond = cond & " AND sale_date <= #" & Format(up_to_date, "mm/dd/yyyy") & "#"
    
    Debug.Print "  [getReturnInwardsQuantity] Condition: " & cond
    result = Nz(DSum("quantity", "qry_return_inwards", cond), 0)
    Debug.Print "  [getReturnInwardsQuantity] Result: " & result
    
    getReturnInwardsQuantity = result
End Function

Public Function getReturnOutwardsQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim cond As String: cond = "1=1"
    Dim result As Long
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        cond = cond & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        cond = cond & " AND product_spec_id = " & product_spec_id
    End If
    If IsDate(up_to_date) Then cond = cond & " AND sale_date <= #" & Format(up_to_date, "mm/dd/yyyy") & "#"
    
    Debug.Print "  [getReturnOutwardsQuantity] Condition: " & cond
    result = Nz(DSum("quantity", "qry_return_outwards", cond), 0)
    Debug.Print "  [getReturnOutwardsQuantity] Result: " & result
    
    getReturnOutwardsQuantity = result
End Function

Public Function getCreditSalesQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim cond As String: cond = "1=1"
    Dim result As Long
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        cond = cond & " AND product_id = " & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        cond = cond & " AND product_spec_id = " & product_spec_id
    End If
    If IsDate(up_to_date) Then cond = cond & " AND sale_date <= #" & Format(up_to_date, "mm/dd/yyyy") & "#"
    
    Debug.Print "  [getCreditSalesQuantity] Condition: " & cond
    result = Nz(DSum("quantity", "tbl_debts", cond), 0)
    Debug.Print "  [getCreditSalesQuantity] Result: " & result
    
    getCreditSalesQuantity = result
End Function

Public Function getCurrentStockQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    Dim qtyPurchased As Long
    Dim qtyReturnedInwards As Long
    Dim qtySold As Long
    Dim qtyCreditSales As Long
    Dim qtyReturnedOutwards As Long
    Dim qtyOfficeUse As Long
    Dim qtyDrawings As Long
    Dim result As Long
    
    Debug.Print "=== getCurrentStockQuantity DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "up_to_date: " & up_to_date
    
    If Not IsDate(up_to_date) Then up_to_date = Date
    Debug.Print "Final up_to_date: " & Format(up_to_date, "mm/dd/yyyy")
    
    ' Stock calculation includes all inventory movements:
    ' INFLOWS: Purchases + Return Inwards (sales returns - inventory increases)
    ' OUTFLOWS: Sales (cash) + Credit Sales + Return Outwards + Office Use + Drawings
    
    qtyPurchased = getPurchasedQuantity(product_spec_id, product_id, up_to_date)
    Debug.Print "Purchased Quantity: " & qtyPurchased
    
    qtyReturnedInwards = getReturnInwardsQuantity(product_spec_id, product_id, up_to_date)
    Debug.Print "Returned Inwards Quantity: " & qtyReturnedInwards
    
    qtySold = getSoldQuantity(product_spec_id, product_id, up_to_date)
    Debug.Print "Sold Quantity: " & qtySold
    
    qtyCreditSales = getCreditSalesQuantity(product_spec_id, product_id, up_to_date)
    Debug.Print "Credit Sales Quantity: " & qtyCreditSales
    
    qtyReturnedOutwards = getReturnOutwardsQuantity(product_spec_id, product_id, up_to_date)
    Debug.Print "Returned Outwards Quantity: " & qtyReturnedOutwards
    
    qtyOfficeUse = getOfficeUseQuantity(product_spec_id, product_id, up_to_date)
    Debug.Print "Office Use Quantity: " & qtyOfficeUse
    
    qtyDrawings = getDrawingsQuantity(product_spec_id, product_id, up_to_date)
    Debug.Print "Drawings Quantity: " & qtyDrawings
    
    result = qtyPurchased + qtyReturnedInwards - qtySold - qtyCreditSales - qtyReturnedOutwards - qtyOfficeUse - qtyDrawings
    Debug.Print "Current Stock = " & qtyPurchased & " + " & qtyReturnedInwards & " - " & qtySold & " - " & qtyCreditSales & " - " & qtyReturnedOutwards & " - " & qtyOfficeUse & " - " & qtyDrawings & " = " & result
    Debug.Print "=== END getCurrentStockQuantity DEBUG ==="
    
    getCurrentStockQuantity = result
End Function

