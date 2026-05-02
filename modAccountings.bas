Attribute VB_Name = "modAccountings"
Option Compare Database
Option Explicit

' ==============================================================================
' MODULE: modAccountings
' PURPOSE: Professional inventory valuation and COGS calculation per GAAP
' METHOD: Weighted Average Costing with comprehensive transaction support
' NAMING CONVENTION:
'   - All functions accept Optional product_spec_id parameter
'   - If product_spec_id provided: returns data for that specific product
'   - If product_spec_id omitted (0): returns aggregated data for all products
'   - Quantity functions: get[Transaction]Qty() - e.g., getPurchasedQty()
'   - Value functions: get[Transaction]() - e.g., getPurchased()
'
' ACCOUNTING INTEGRATION - HOW TRANSACTIONS AFFECT FINANCIAL STATEMENTS:
'
' INCOME STATEMENT:
'   1. Net Purchases = Purchases - Return Outwards
'      (Purchase Returns reduce the cost of goods acquired)
'
'   2. Net Sales = Direct Sales + Credit Sales - Return Inwards
'      (Sales Returns reduce total revenue)
'
'   3. COGS = (Sold + Credit Sales + Office Use + Drawings) × Weighted Avg Cost
'      (All inventory outflows are costed at weighted average purchase price)
'
'   4. Gross Profit = Net Sales - COGS
'
' BALANCE SHEET:
'   Opening Stock = Opening Qty × Weighted Avg Cost (as of period start)
'   Closing Stock = Closing Qty × Weighted Avg Cost (as of period end)
'
' TRANSACTION HANDLING - DETAILED ACCOUNTING IMPACT:
'
' 1. RETURNS INWARDS (Sales Returns):
'    - Stock Impact: + Increases inventory (goods come back)
'    - Net Sales: - Reduces Net Sales (reverses original sale revenue)
'    - Net Purchases: No impact
'    - COGS: NOT included (reverses original sale's COGS automatically)
'    - Gross Profit: Indirectly increases (Net Sales decreases, but original COGS was reversed)
'
' 2. RETURNS OUTWARDS (Purchase Returns):
'    - Stock Impact: - Decreases inventory (goods sent back to supplier)
'    - Net Sales: No impact
'    - Net Purchases: - Reduces Net Purchases (reverses original purchase cost)
'    - COGS: NOT included (affects weighted average cost calculation)
'    - Gross Profit: Indirectly increases (Net Purchases decreases, reducing cost basis)
'
' 3. DRAWINGS (Owner Withdrawals):
'    - Stock Impact: - Decreases inventory (goods withdrawn by owner)
'    - Net Sales: No impact (not a sale)
'    - Net Purchases: No impact
'    - COGS: + Included (goods left business, costed at weighted average)
'    - Gross Profit: Decreases (COGS increases, but no corresponding revenue)
'    - Equity: Reduces owner's equity (drawings account)
'
' 4. OFFICE USE (Internal Consumption):
'    - Stock Impact: - Decreases inventory (goods consumed internally)
'    - Net Sales: No impact (not a sale)
'    - Net Purchases: No impact
'    - COGS: + Included (goods left inventory, costed at weighted average)
'    - Gross Profit: Decreases (COGS increases, but no corresponding revenue)
'    - Expense: Expensed as operating cost
'
' 5. CREDIT SALES (Debts/Accounts Receivable):
'    - Stock Impact: - Decreases inventory (same as direct sales)
'    - Net Sales: + Included (revenue recognized when goods sold, not when paid)
'    - Net Purchases: No impact
'    - COGS: + Included (goods left inventory, costed at weighted average)
'    - Gross Profit: Calculated normally (Net Sales - COGS)
'    - Receivables: Creates accounts receivable asset
'
' VALIDATION EQUATION:
'   Opening Stock + Net Purchases = COGS + Closing Stock
'   (Must always balance - validates inventory accounting integrity)
' ==============================================================================

' ==============================================================================
' SECTION 1: QUANTITY FUNCTIONS
' These track physical movement of inventory
' ==============================================================================

' Purchases - goods acquired from suppliers (INCREASES inventory)
' If product_spec_id = 0: returns total for all products
' If product_spec_id > 0: returns total for specific product
Public Function getPurchasedQty(Optional product_spec_id As Long = 0, _
                                Optional product_id As Long = 0, _
                                Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "", _
                                Optional upToDate As Variant = "") As Double
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getPurchasedQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getPurchasedQty = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(purchase_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(purchase_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(purchase_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getPurchasedQty = Nz(DSum("quantity", "qry_purchases", condition), 0)
End Function

' Sales - goods sold to customers (DECREASES inventory)
Public Function getSoldQty(Optional product_spec_id As Long = 0, _
                           Optional product_id As Long = 0, _
                           Optional month_id As Long = 0, _
                           Optional startDate As Variant = "", _
                           Optional EndDate As Variant = "", _
                           Optional upToDate As Variant = "") As Double
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getSoldQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getSoldQty = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getSoldQty = Nz(DSum("quantity", "tbl_sales", condition), 0)
End Function

' Sales Returns (Return Inwards) - defective goods returned by customers (INCREASES inventory)
Public Function getReturnedInwardsQty(Optional product_spec_id As Long = 0, _
                                      Optional product_id As Long = 0, _
                                      Optional month_id As Long = 0, _
                                      Optional startDate As Variant = "", _
                                      Optional EndDate As Variant = "", _
                                      Optional upToDate As Variant = "") As Double
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getReturnedInwardsQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getReturnedInwardsQty = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getReturnedInwardsQty = Nz(DSum("quantity", "qry_return_inwards", condition), 0)
End Function

' Purchase Returns (Return Outwards) - defective goods returned to suppliers (DECREASES inventory)
Public Function getReturnedOutwardsQty(Optional product_spec_id As Long = 0, _
                                       Optional product_id As Long = 0, _
                                       Optional month_id As Long = 0, _
                                       Optional startDate As Variant = "", _
                                       Optional EndDate As Variant = "", _
                                       Optional upToDate As Variant = "") As Double
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getReturnedOutwardsQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getReturnedOutwardsQty = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getReturnedOutwardsQty = Nz(DSum("quantity", "qry_return_outwards", condition), 0)
End Function

' Office Use - goods consumed for internal operations (DECREASES inventory)
Public Function getOfficeUseQty(Optional product_spec_id As Long = 0, _
                                Optional product_id As Long = 0, _
                                Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "", _
                                Optional upToDate As Variant = "") As Double
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getOfficeUseQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getOfficeUseQty = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getOfficeUseQty = Nz(DSum("quantity", "qry_sales_office_use", condition), 0)
End Function

' Drawings - goods withdrawn by owner for personal use (DECREASES inventory)
Public Function getDrawingsQty(Optional product_spec_id As Long = 0, _
                               Optional product_id As Long = 0, _
                               Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional EndDate As Variant = "", _
                               Optional upToDate As Variant = "") As Double
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getDrawingsQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getDrawingsQty = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getDrawingsQty = Nz(DSum("quantity", "qry_drawings", condition), 0)
End Function

' Credit Sales (Accounts Receivable) - goods sold on credit terms (DECREASES inventory)
Public Function getCreditSalesQty(Optional product_spec_id As Long = 0, _
                                  Optional product_id As Long = 0, _
                                  Optional month_id As Long = 0, _
                                  Optional startDate As Variant = "", _
                                  Optional EndDate As Variant = "", _
                                  Optional upToDate As Variant = "") As Double
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getCreditSalesQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getCreditSalesQty = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getCreditSalesQty = Nz(DSum("quantity", "tbl_debts", condition), 0)
End Function

' ==============================================================================
' SECTION 2: VALUE FUNCTIONS (MONETARY)
' These track monetary amounts - no "Value" suffix needed
' ==============================================================================

' Total purchase cost paid to suppliers
Public Function getPurchased(Optional product_spec_id As Long = 0, _
                             Optional product_id As Long = 0, _
                             Optional month_id As Long = 0, _
                             Optional startDate As Variant = "", _
                             Optional EndDate As Variant = "", _
                             Optional upToDate As Variant = "") As Currency
    Dim condition As String
    Dim result As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getPurchased(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getPurchased = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(purchase_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getPurchased] Using upToDate filter: <= " & Format(CDate(upToDate), "mm/dd/yyyy")
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(purchase_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(purchase_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getPurchased] Using date range: " & Format(CDate(startDate), "mm/dd/yyyy") & " to " & Format(CDate(EndDate), "mm/dd/yyyy")
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
        Debug.Print "      [getPurchased] Using month_id: " & month_id
    End If
    
    Debug.Print "      [getPurchased] Condition: " & condition
    result = Nz(DSum("amount", "qry_purchases", condition), 0)
    Debug.Print "      [getPurchased] Result: " & Format(result, "#,##0.00")
    
    getPurchased = result
End Function

' Total direct sales revenue from customers
Public Function getDirectSales(Optional product_spec_id As Long = 0, _
                               Optional product_id As Long = 0, _
                               Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional EndDate As Variant = "", _
                               Optional upToDate As Variant = "") As Currency
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getDirectSales(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getDirectSales = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getDirectSales = Nz(DSum("amount", "qry_sales", condition), 0)
End Function

' Credit sales revenue (accounts receivable created)
Public Function getCreditSales(Optional product_spec_id As Long = 0, _
                               Optional product_id As Long = 0, _
                               Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional EndDate As Variant = "", _
                               Optional upToDate As Variant = "") As Currency
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getCreditSales(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getCreditSales = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getCreditSales = Nz(DSum("amount", "tbl_debts", condition), 0)
End Function

' Total sales revenue (Direct Sales + Credit Sales)
Public Function getSales(Optional product_spec_id As Long = 0, _
                         Optional product_id As Long = 0, _
                         Optional month_id As Long = 0, _
                         Optional startDate As Variant = "", _
                         Optional EndDate As Variant = "", _
                         Optional upToDate As Variant = "") As Currency
    Dim directSales As Currency
    Dim creditSales As Currency
    
    directSales = getDirectSales(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    creditSales = getCreditSales(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    
    getSales = directSales + creditSales
End Function

' Sales returns value (refunds/credits to customers)
Public Function getReturnedInwards(Optional product_spec_id As Long = 0, _
                                   Optional product_id As Long = 0, _
                                   Optional month_id As Long = 0, _
                                   Optional startDate As Variant = "", _
                                   Optional EndDate As Variant = "", _
                                   Optional upToDate As Variant = "") As Currency
    Dim condition As String
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getReturnedInwards(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getReturnedInwards = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getReturnedInwards = Nz(DSum("amount", "qry_return_inwards", condition), 0)
End Function

' Purchase returns value (refunds/credits from suppliers)
Public Function getReturnedOutwards(Optional product_spec_id As Long = 0, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional EndDate As Variant = "", _
                                    Optional upToDate As Variant = "") As Currency
    Dim condition As String
    Dim result As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getReturnedOutwards(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getReturnedOutwards = total
        Exit Function
    End If
    
    ' Single product logic
    If product_id > 0 Then
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getReturnedOutwards] Using upToDate filter: <= " & Format(CDate(upToDate), "mm/dd/yyyy")
    ElseIf IsDate(startDate) And IsDate(EndDate) Then
        condition = condition & " AND datevalue(sale_date)>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND datevalue(sale_date)<=#" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getReturnedOutwards] Using date range: " & Format(CDate(startDate), "mm/dd/yyyy") & " to " & Format(CDate(EndDate), "mm/dd/yyyy")
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
        Debug.Print "      [getReturnedOutwards] Using month_id: " & month_id
    End If
    
    Debug.Print "      [getReturnedOutwards] Condition: " & condition
    result = Nz(DSum("amount", "qry_return_outwards", condition), 0)
    Debug.Print "      [getReturnedOutwards] Result: " & Format(result, "#,##0.00")
    
    getReturnedOutwards = result
End Function

' ==============================================================================
' SECTION 3: NET INVENTORY POSITION
' Calculates actual quantity on hand considering ALL transactions
' ==============================================================================

' Net inventory quantity = Purchases + ReturnInwards - Sales - CreditSales - ReturnOutwards - OfficeUse - Drawings
Public Function getInventoryQty(Optional product_spec_id As Long = 0, _
                                Optional product_id As Long = 0, _
                                Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "", _
                                Optional upToDate As Variant = "") As Double
    Dim netQty As Double
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getInventoryQty(CLng(productID), 0, month_id, startDate, EndDate, upToDate)
        Next productID
        
        getInventoryQty = total
        Exit Function
    End If
    
    ' Single product logic
    'Debug.Print "    [getInventoryQty] product_spec_id: " & product_spec_id & ", product_id: " & product_id
    'Debug.Print "    [getInventoryQty] month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & endDate & ", upToDate: " & upToDate
    
    ' INFLOWS (goods coming IN to inventory)
    netQty = getPurchasedQty(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    'Debug.Print "    [getInventoryQty] Purchased: " & netQty
    
    netQty = netQty + getReturnedInwardsQty(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    'Debug.Print "    [getInventoryQty] After Return Inwards: " & netQty
    
    ' OUTFLOWS (goods going OUT of inventory)
    netQty = netQty - getSoldQty(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    'Debug.Print "    [getInventoryQty] After Sales: " & netQty
    
    netQty = netQty - getCreditSalesQty(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    'Debug.Print "    [getInventoryQty] After Credit Sales: " & netQty
    
    netQty = netQty - getReturnedOutwardsQty(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    'Debug.Print "    [getInventoryQty] After Return Outwards: " & netQty
    
    netQty = netQty - getOfficeUseQty(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    'Debug.Print "    [getInventoryQty] After Office Use: " & netQty
    
    netQty = netQty - getDrawingsQty(product_spec_id, product_id, month_id, startDate, EndDate, upToDate)
    'Debug.Print "    [getInventoryQty] Final Net Quantity: " & netQty
    
    getInventoryQty = netQty
End Function

' ==============================================================================
' SECTION 4: WEIGHTED AVERAGE COST
' Core costing method - ensures time-consistent historical valuations
' ==============================================================================

' Calculate weighted average cost per unit based on cumulative purchases up to date
' NOTE: Only applicable for single product (product_spec_id must be provided)
Public Function getWeightedAvgCost(product_spec_id As Long, upToDate As Date, Optional product_id As Long = 0) As Double
    Dim totalCost As Currency
    Dim totalQty As Double
    Dim result As Double
    
    Debug.Print "  [getWeightedAvgCost] product_spec_id: " & product_spec_id & ", product_id: " & product_id & ", upToDate: " & Format(upToDate, "mm/dd/yyyy")
    
    totalCost = getPurchased(product_spec_id, product_id, , , , upToDate)
    Debug.Print "  [getWeightedAvgCost] Total Purchase Cost up to " & Format(upToDate, "mm/dd/yyyy") & ": " & Format(totalCost, "#,##0.00")
    
    totalQty = getPurchasedQty(product_spec_id, product_id, , , , upToDate)
    Debug.Print "  [getWeightedAvgCost] Total Purchase Quantity up to " & Format(upToDate, "mm/dd/yyyy") & ": " & totalQty
    
    If totalQty = 0 Then
        Debug.Print "  [getWeightedAvgCost] TotalQty = 0, returning 0"
        getWeightedAvgCost = 0
    Else
        result = totalCost / totalQty
        Debug.Print "  [getWeightedAvgCost] Weighted Avg Cost (TotalCost/TotalQty): " & Format(result, "#,##0.00")
        getWeightedAvgCost = result
    End If
End Function

' ==============================================================================
' SECTION 5: OPENING & CLOSING STOCK
' Balance sheet inventory positions at period boundaries
' ==============================================================================

' Opening stock QUANTITY at start of period
Public Function getOpeningQty(Optional product_spec_id As Long = 0, _
                              Optional periodStartDate As Variant = Null, _
                              Optional product_id As Long = 0) As Double
    Dim upToDate As Date
    Dim result As Double
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getOpeningQty(CLng(productID), periodStartDate, 0)
        Next productID
        
        getOpeningQty = total
        Exit Function
    End If
    
    ' Single product logic
    If IsNull(periodStartDate) Or Not IsDate(periodStartDate) Then
        getOpeningQty = 0
        Exit Function
    End If
    
    upToDate = CDate(periodStartDate) - 1
    Debug.Print "  [getOpeningQty] periodStartDate: " & Format(CDate(periodStartDate), "mm/dd/yyyy") & ", upToDate: " & Format(upToDate, "mm/dd/yyyy")
    Debug.Print "  [getOpeningQty] NOTE: upToDate is one day BEFORE periodStartDate because opening stock = closing stock of previous day"
    result = getInventoryQty(product_spec_id, product_id, , , , upToDate)
    Debug.Print "  [getOpeningQty] Opening Quantity: " & result
    getOpeningQty = result
End Function

' Closing stock QUANTITY at end of period
Public Function getClosingQty(Optional product_spec_id As Long = 0, _
                              Optional periodEndDate As Variant = Null, _
                              Optional product_id As Long = 0) As Double
    Dim result As Double
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Double
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getClosingQty(CLng(productID), periodEndDate, 0)
        Next productID
        
        getClosingQty = total
        Exit Function
    End If
    
    ' Single product logic
    If IsNull(periodEndDate) Or Not IsDate(periodEndDate) Then
        getClosingQty = 0
        Exit Function
    End If
    
    Debug.Print "  [getClosingQty] periodEndDate: " & Format(CDate(periodEndDate), "mm/dd/yyyy")
    result = getInventoryQty(product_spec_id, product_id, , , , CDate(periodEndDate))
    Debug.Print "  [getClosingQty] Closing Quantity: " & result
    getClosingQty = result
End Function

' Opening stock VALUE at start of period (for balance sheet)
Public Function getOpeningStock(Optional product_spec_id As Long = 0, _
                                Optional product_id As Long = 0, _
                                Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "") As Currency
    Dim periodStartDate As Date
    Dim priceDate As Date
    Dim openingQty As Double
    Dim avgCost As Double
    Dim result As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getOpeningStock(CLng(productID), 0, month_id, startDate, EndDate)
        Next productID
        
        getOpeningStock = total
        Exit Function
    End If
    
    ' Single product logic
    Debug.Print "=== getOpeningStock DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & EndDate
    
    ' Determine period start date
    If IsDate(startDate) Then
        periodStartDate = CDate(startDate)
        Debug.Print "Using startDate: " & Format(periodStartDate, "mm/dd/yyyy")
    ElseIf month_id > 0 Then
        periodStartDate = DateSerial(Year(Date), month_id, 1)
        Debug.Print "Using month_id, periodStartDate: " & Format(periodStartDate, "mm/dd/yyyy")
    Else
        Debug.Print "No valid period start date - returning 0"
        getOpeningStock = 0
        Exit Function
    End If
    
    openingQty = getOpeningQty(product_spec_id, periodStartDate, product_id)
    Debug.Print "Opening Quantity: " & openingQty
    
    If openingQty <= 0 Then
        Debug.Print "Opening Quantity <= 0 - returning 0"
        Debug.Print "=== END getOpeningStock DEBUG ==="
        getOpeningStock = 0
        Exit Function
    End If
    
    ' Value at weighted average cost as of day before period starts
    priceDate = periodStartDate - 1
    Debug.Print "Price Date (periodStartDate - 1): " & Format(priceDate, "mm/dd/yyyy")
    avgCost = getWeightedAvgCost(product_spec_id, priceDate, product_id)
    Debug.Print "Weighted Average Cost: " & Format(avgCost, "#,##0.00")
    
    result = CCur(openingQty * avgCost)
    Debug.Print "Opening Stock Value (Qty × AvgCost): " & Format(result, "#,##0.00")
    Debug.Print "=== END getOpeningStock DEBUG ==="
    
    getOpeningStock = result
End Function

' Closing stock VALUE at end of period (for balance sheet)
Public Function getClosingStock(Optional product_spec_id As Long = 0, _
                                Optional product_id As Long = 0, _
                                Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "") As Currency
    Dim periodEndDate As Date
    Dim closingQty As Double
    Dim avgCost As Double
    Dim result As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getClosingStock(CLng(productID), 0, month_id, startDate, EndDate)
        Next productID
        
        getClosingStock = CCur(total)
        Exit Function
    End If
    
    ' Single product logic
    Debug.Print "=== getClosingStock DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & EndDate
    
    ' Determine period end date
    Dim endDateStr As String
    If Not IsNull(EndDate) And Not IsEmpty(EndDate) And CStr(EndDate) <> "" Then
        endDateStr = Trim(CStr(EndDate))
        endDateStr = Replace(endDateStr, "#", "")
        endDateStr = Replace(endDateStr, """", "")
        If IsDate(endDateStr) Then
            periodEndDate = CDate(endDateStr)
            Debug.Print "Using endDate: " & Format(periodEndDate, "mm/dd/yyyy")
        ElseIf month_id > 0 Then
            periodEndDate = DateSerial(Year(Date), month_id + 1, 0)
            Debug.Print "endDate invalid, using month_id: " & Format(periodEndDate, "mm/dd/yyyy")
        Else
            periodEndDate = Date
            Debug.Print "endDate invalid, using current date: " & Format(periodEndDate, "mm/dd/yyyy")
        End If
    ElseIf month_id > 0 Then
        periodEndDate = DateSerial(Year(Date), month_id + 1, 0)
        Debug.Print "Using month_id: " & Format(periodEndDate, "mm/dd/yyyy")
    Else
        periodEndDate = Date
        Debug.Print "Using current date: " & Format(periodEndDate, "mm/dd/yyyy")
    End If
    
    closingQty = getClosingQty(product_spec_id, periodEndDate, product_id)
    Debug.Print "Closing Quantity: " & closingQty
    
    If closingQty <= 0 Then
        Debug.Print "Closing Quantity <= 0 - returning 0"
        Debug.Print "=== END getClosingStock DEBUG ==="
        getClosingStock = 0
        Exit Function
    End If
    
    avgCost = getWeightedAvgCost(product_spec_id, periodEndDate, product_id)
    Debug.Print "Weighted Average Cost: " & Format(avgCost, "#,##0.00")
    
    result = CCur(closingQty * avgCost)
    Debug.Print "Closing Stock Value: " & Format(result, "#,##0.00")
    Debug.Print "=== END getClosingStock DEBUG ==="
    
    getClosingStock = result
End Function

' ==============================================================================
' SECTION 6: COGS (Cost of Goods Sold)
' The cost of ALL goods that left inventory
' ==============================================================================

' Calculate COGS: Opening Stock + Net Purchases - Closing Stock
Public Function getCOGS(Optional product_spec_id As Long = 0, _
                        Optional product_id As Long = 0, _
                        Optional month_id As Long = 0, _
                        Optional startDate As Variant = "", _
                        Optional EndDate As Variant = "") As Currency
    Dim openingStock As Currency
    Dim netPurchases As Currency
    Dim closingStock As Currency
    Dim cogs As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getCOGS(CLng(productID), 0, month_id, startDate, EndDate)
        Next productID
        
        getCOGS = total
        Exit Function
    End If
    
    ' Single product logic
    Debug.Print "=== getCOGS DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    
    openingStock = getOpeningStock(product_spec_id, product_id, month_id, startDate, EndDate)
    Debug.Print "Opening Stock: " & Format(openingStock, "#,##0.00")
    
    netPurchases = getNetPurchases(product_spec_id, product_id, month_id, startDate, EndDate)
    Debug.Print "Net Purchases: " & Format(netPurchases, "#,##0.00")
    
    closingStock = getClosingStock(product_spec_id, product_id, month_id, startDate, EndDate)
    Debug.Print "Closing Stock: " & Format(closingStock, "#,##0.00")
    
    cogs = openingStock + netPurchases - closingStock
    Debug.Print "COGS: " & Format(cogs, "#,##0.00")
    Debug.Print "=== END getCOGS DEBUG ==="
    
    getCOGS = cogs
End Function

' ==============================================================================
' SECTION 7: ACCOUNTING CONCEPTS (GAAP)
' Professional accounting calculations
' ==============================================================================

' Net Purchases = Purchases - Purchase Returns
Public Function getNetPurchases(Optional product_spec_id As Long = 0, _
                                Optional product_id As Long = 0, _
                                Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional EndDate As Variant = "") As Currency
    Dim purchases As Currency
    Dim returnOut As Currency
    Dim result As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getNetPurchases(CLng(productID), 0, month_id, startDate, EndDate)
        Next productID
        
        getNetPurchases = total
        Exit Function
    End If
    
    ' Single product logic
    Debug.Print "=== getNetPurchases DEBUG ==="
    
    purchases = getPurchased(product_spec_id, product_id, month_id, startDate, EndDate)
    Debug.Print "Purchases: " & Format(purchases, "#,##0.00")
    
    returnOut = getReturnedOutwards(product_spec_id, product_id, month_id, startDate, EndDate)
    Debug.Print "Return Outwards: " & Format(returnOut, "#,##0.00")
    
    result = purchases - returnOut
    Debug.Print "Net Purchases: " & Format(result, "#,##0.00")
    Debug.Print "=== END getNetPurchases DEBUG ==="
    
    getNetPurchases = result
End Function

' Net Sales = Direct Sales + Credit Sales - Sales Returns
Public Function getNetSales(Optional product_spec_id As Long = 0, _
                            Optional product_id As Long = 0, _
                            Optional month_id As Long = 0, _
                            Optional startDate As Variant = "", _
                            Optional EndDate As Variant = "") As Currency
    Dim directSales As Currency
    Dim creditSales As Currency
    Dim returnIn As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getNetSales(CLng(productID), 0, month_id, startDate, EndDate)
        Next productID
        
        getNetSales = total
        Exit Function
    End If
    
    ' Single product logic
    directSales = getDirectSales(product_spec_id, product_id, month_id, startDate, EndDate)
    creditSales = getCreditSales(product_spec_id, product_id, month_id, startDate, EndDate)
    returnIn = getReturnedInwards(product_spec_id, product_id, month_id, startDate, EndDate)
    
    Debug.Print "Direct Sales: " & Format(directSales, "#,##0.00")
    Debug.Print "Credit Sales: " & Format(creditSales, "#,##0.00")
    Debug.Print "Return Inwards: " & Format(returnIn, "#,##0.00")
    
    getNetSales = directSales + creditSales - returnIn
End Function

' Gross Profit = Net Sales - COGS
Public Function getGrossProfit(Optional product_spec_id As Long = 0, _
                               Optional product_id As Long = 0, _
                               Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional EndDate As Variant = "") As Currency
    Dim netSales As Currency
    Dim cogs As Currency
    
    netSales = getNetSales(product_spec_id, product_id, month_id, startDate, EndDate)
    cogs = getCOGS(product_spec_id, product_id, month_id, startDate, EndDate)
    
    getGrossProfit = netSales - cogs
End Function

' Gross Profit Margin (percentage)
Public Function getGrossMargin(Optional product_spec_id As Long = 0, _
                               Optional product_id As Long = 0, _
                               Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional EndDate As Variant = "") As Double
    Dim netSales As Currency
    Dim grossProfit As Currency
    
    netSales = getNetSales(product_spec_id, product_id, month_id, startDate, EndDate)
    
    If netSales = 0 Then
        getGrossMargin = 0
    Else
        grossProfit = getGrossProfit(product_spec_id, product_id, month_id, startDate, EndDate)
        getGrossMargin = Round((grossProfit / netSales) * 100, 2)
    End If
End Function

' Cost of Goods Available for Sale = Opening Stock + Net Purchases
Public Function getCOGAS(Optional product_spec_id As Long = 0, _
                         Optional product_id As Long = 0, _
                         Optional month_id As Long = 0, _
                         Optional startDate As Variant = "", _
                         Optional EndDate As Variant = "") As Currency
    Dim openingStock As Currency
    Dim netPurchases As Currency
    
    If product_spec_id = 0 And product_id = 0 Then
        ' Aggregate for all products
        Dim products As Collection
        Dim productID As Variant
        Dim total As Currency
        
        total = 0
        Set products = getAllProductSpecIDs()
        
        For Each productID In products
            total = total + getCOGAS(CLng(productID), 0, month_id, startDate, EndDate)
        Next productID
        
        getCOGAS = total
        Exit Function
    End If
    
    ' Single product logic
    openingStock = getOpeningStock(product_spec_id, product_id, month_id, startDate, EndDate)
    netPurchases = getNetPurchases(product_spec_id, product_id, month_id, startDate, EndDate)
    
    getCOGAS = openingStock + netPurchases
End Function

' ==============================================================================
' SECTION 8: UTILITY FUNCTIONS
' ==============================================================================

' Get list of all product_spec_ids that have ever been purchased
Private Function getAllProductSpecIDs() As Collection
    Dim rs As DAO.Recordset
    Dim col As New Collection
    Dim productID As Long
    
    Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT product_spec_id FROM qry_purchases ORDER BY product_spec_id")
    
    Do While Not rs.EOF
        productID = rs!product_spec_id
        col.Add productID
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set getAllProductSpecIDs = col
End Function

' ==============================================================================
' SECTION 9: VALIDATION & REPORTING
' ==============================================================================

' Validate inventory equation: Opening + Net Purchases = COGS + Closing
Public Function ValidateInventoryEquation(Optional product_spec_id As Long = 0, _
                                          Optional product_id As Long = 0, _
                                          Optional month_id As Long = 0, _
                                          Optional startDate As Variant = "", _
                                          Optional EndDate As Variant = "") As Boolean
    Dim openingStock As Currency
    Dim netPurchases As Currency
    Dim cogs As Currency
    Dim closingStock As Currency
    Dim leftSide As Currency
    Dim rightSide As Currency
    Dim difference As Currency
    
    Debug.Print vbCrLf & "========================================"
    Debug.Print "INVENTORY EQUATION VALIDATION"
    Debug.Print "========================================"
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "----------------------------------------"
    
    openingStock = getOpeningStock(product_spec_id, product_id, month_id, startDate, EndDate)
    netPurchases = getNetPurchases(product_spec_id, product_id, month_id, startDate, EndDate)
    cogs = getCOGS(product_spec_id, product_id, month_id, startDate, EndDate)
    closingStock = getClosingStock(product_spec_id, product_id, month_id, startDate, EndDate)
    
    leftSide = openingStock + netPurchases
    rightSide = cogs + closingStock
    difference = Abs(leftSide - rightSide)
    
    Debug.Print "Opening Stock: " & Format(openingStock, "#,##0.00")
    Debug.Print "Net Purchases: " & Format(netPurchases, "#,##0.00")
    Debug.Print "COGS: " & Format(cogs, "#,##0.00")
    Debug.Print "Closing Stock: " & Format(closingStock, "#,##0.00")
    Debug.Print "Left (Opening + Net): " & Format(leftSide, "#,##0.00")
    Debug.Print "Right (COGS + Closing): " & Format(rightSide, "#,##0.00")
    Debug.Print "Difference: " & Format(difference, "#,##0.00")
    Debug.Print "Status: " & IIf(difference < 0.01, "PASSED", "FAILED")
    Debug.Print "========================================" & vbCrLf
    
    ValidateInventoryEquation = (difference < 0.01)
End Function

' ==============================================================================
' SECTION 10: BACKWARD COMPATIBILITY
' Legacy function names for existing queries
' ==============================================================================

' Backward compatibility wrapper for getCurrentStockQuantity
' Returns inventory quantity - aggregates all products if product_spec_id = 0
Public Function getCurrentStockQuantity( _
    Optional ByVal product_spec_id As Long = 0, _
    Optional ByVal product_id As Long = 0, _
    Optional ByVal up_to_date As Variant = Null) As Long
    
    getCurrentStockQuantity = getInventoryQty(product_spec_id, product_id, 0, 0, 0, up_to_date)
End Function


Public Function getNetProfit(Optional startDate As Variant = "", Optional EndDate As Variant = "") As Currency
    Dim grossProfit As Currency
    Dim expenses As Currency
    
    grossProfit = getGrossProfit(0, 0, 0, startDate, EndDate)
    expenses = getExpenses(0, 0, startDate, EndDate)
    Debug.Print "Gross Profit: " & grossProfit
    Debug.Print "Expenses    : " & expenses
    getNetProfit = grossProfit - expenses
End Function




Public Function getOfficeUseSales(Optional startDate As Variant = "", Optional EndDate As Variant = "") As Currency
    Dim tableSQL As String
    Dim amount As Currency
    Dim criteria As String

    tableSQL = "qry_sales_office_use"

    ' Build criteria safely
    If IsDate(startDate) Then
        criteria = criteria & " AND DateValue(sale_date) >= #" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
    End If

    If IsDate(EndDate) Then
        criteria = criteria & " AND DateValue(sale_date) <= #" & Format(CDate(EndDate), "mm/dd/yyyy") & "#"
    End If

    ' Remove leading AND if exists
    If Len(criteria) > 0 Then
        criteria = Mid(criteria, 6)
    End If

    ' Get sum
    amount = Nz(DSum("amount", tableSQL, criteria), 0)

    ' Return value
    getOfficeUseSales = amount
End Function

