Option Compare Database
Option Explicit

' ==============================================================================
' MODULE: modAccountings
' PURPOSE: Professional inventory valuation and COGS calculation per GAAP
' METHOD: Weighted Average Costing with comprehensive transaction support
' NAMING CONVENTION:
'   - getItem*() = Single product (product_spec_id required)
'   - get*()     = All products aggregated (iterates over all product_spec_id)
'
' ACCOUNTING INTEGRATION - HOW TRANSACTIONS AFFECT FINANCIAL STATEMENTS:
'
' INCOME STATEMENT:
'   1. Net Purchases = Purchases - Return Outwards
'      (Purchase Returns reduce the cost of goods acquired)
'
'   2. Net Sales = Cash Sales + Credit Sales - Return Inwards
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
'    - Stock Impact: - Decreases inventory (same as cash sales)
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
' SECTION 1: ITEM-LEVEL QUANTITY FUNCTIONS (SINGLE PRODUCT)
' These track physical movement for ONE product_spec_id
' ==============================================================================

' Purchases - goods acquired from suppliers (INCREASES inventory)
Public Function getItemQtyPurchased(product_spec_id As Long, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional endDate As Variant = "", _
                                    Optional upToDate As Variant = "") As Double
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND purchase_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND purchase_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND purchase_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemQtyPurchased = Nz(DSum("quantity", "qry_purchases", condition), 0)
End Function

' Sales - goods sold to customers (DECREASES inventory)
Public Function getItemQtySold(product_spec_id As Long, _
                               Optional product_id As Long = 0, _
                               Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional endDate As Variant = "", _
                               Optional upToDate As Variant = "") As Double
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemQtySold = Nz(DSum("quantity", "tbl_sales", condition), 0)
End Function

' Sales Returns (Return Inwards) - defective goods returned by customers (INCREASES inventory)
Public Function getItemQtyReturnedInwards(product_spec_id As Long, _
                                          Optional product_id As Long = 0, _
                                          Optional month_id As Long = 0, _
                                          Optional startDate As Variant = "", _
                                          Optional endDate As Variant = "", _
                                          Optional upToDate As Variant = "") As Double
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemQtyReturnedInwards = Nz(DSum("quantity", "qry_return_inwards", condition), 0)
End Function

' Purchase Returns (Return Outwards) - defective goods returned to suppliers (DECREASES inventory)
Public Function getItemQtyReturnedOutwards(product_spec_id As Long, _
                                           Optional product_id As Long = 0, _
                                           Optional month_id As Long = 0, _
                                           Optional startDate As Variant = "", _
                                           Optional endDate As Variant = "", _
                                           Optional upToDate As Variant = "") As Double
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemQtyReturnedOutwards = Nz(DSum("quantity", "qry_return_outwards", condition), 0)
End Function

' Office Use - goods consumed for internal operations (DECREASES inventory)
Public Function getItemQtyOfficeUse(product_spec_id As Long, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional endDate As Variant = "", _
                                    Optional upToDate As Variant = "") As Double
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemQtyOfficeUse = Nz(DSum("quantity", "qry_sales_office_use", condition), 0)
End Function

' Drawings - goods withdrawn by owner for personal use (DECREASES inventory)
Public Function getItemQtyDrawings(product_spec_id As Long, _
                                   Optional product_id As Long = 0, _
                                   Optional month_id As Long = 0, _
                                   Optional startDate As Variant = "", _
                                   Optional endDate As Variant = "", _
                                   Optional upToDate As Variant = "") As Double
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemQtyDrawings = Nz(DSum("quantity", "qry_drawings", condition), 0)
End Function

' Credit Sales (Accounts Receivable) - goods sold on credit terms (DECREASES inventory)
Public Function getItemQtyCreditSales(product_spec_id As Long, _
                                      Optional product_id As Long = 0, _
                                      Optional month_id As Long = 0, _
                                      Optional startDate As Variant = "", _
                                      Optional endDate As Variant = "", _
                                      Optional upToDate As Variant = "") As Double
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemQtyCreditSales = Nz(DSum("quantity", "tbl_debts", condition), 0)
End Function

' ==============================================================================
' SECTION 2: ITEM-LEVEL VALUE FUNCTIONS (SINGLE PRODUCT - MONETARY)
' These track monetary amounts for ONE product_spec_id
' ==============================================================================

' Total purchase cost paid to suppliers
Public Function getItemPurchaseValue(product_spec_id As Long, _
                                     Optional product_id As Long = 0, _
                                     Optional month_id As Long = 0, _
                                     Optional startDate As Variant = "", _
                                     Optional endDate As Variant = "", _
                                     Optional upToDate As Variant = "") As Currency
    Dim condition As String
    Dim result As Currency
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND purchase_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getItemPurchaseValue] Using upToDate filter: <= " & Format(CDate(upToDate), "mm/dd/yyyy")
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND purchase_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND purchase_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getItemPurchaseValue] Using date range: " & Format(CDate(startDate), "mm/dd/yyyy") & " to " & Format(CDate(endDate), "mm/dd/yyyy")
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
        Debug.Print "      [getItemPurchaseValue] Using month_id: " & month_id
    End If
    
    Debug.Print "      [getItemPurchaseValue] Condition: " & condition
    result = Nz(DSum("amount", "qry_purchases", condition), 0)
    Debug.Print "      [getItemPurchaseValue] Result: " & Format(result, "#,##0.00")
    
    getItemPurchaseValue = result
End Function

' Total sales revenue from customers (cash sales only)
Public Function getItemSalesRevenue(product_spec_id As Long, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional endDate As Variant = "", _
                                    Optional upToDate As Variant = "") As Currency
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemSalesRevenue = Nz(DSum("amount", "qry_sales", condition), 0)
End Function

' Credit sales revenue (accounts receivable created)
Public Function getItemCreditSalesRevenue(product_spec_id As Long, _
                                          Optional product_id As Long = 0, _
                                          Optional month_id As Long = 0, _
                                          Optional startDate As Variant = "", _
                                          Optional endDate As Variant = "", _
                                          Optional upToDate As Variant = "") As Currency
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemCreditSalesRevenue = Nz(DSum("amount", "tbl_debts", condition), 0)
End Function

' Sales returns value (refunds/credits to customers)
Public Function getItemReturnInwardsValue(product_spec_id As Long, _
                                          Optional product_id As Long = 0, _
                                          Optional month_id As Long = 0, _
                                          Optional startDate As Variant = "", _
                                          Optional endDate As Variant = "", _
                                          Optional upToDate As Variant = "") As Currency
    Dim condition As String
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
    End If
    
    getItemReturnInwardsValue = Nz(DSum("amount", "qry_return_inwards", condition), 0)
End Function

' Purchase returns value (refunds/credits from suppliers)
Public Function getItemReturnOutwardsValue(product_spec_id As Long, _
                                           Optional product_id As Long = 0, _
                                           Optional month_id As Long = 0, _
                                           Optional startDate As Variant = "", _
                                           Optional endDate As Variant = "", _
                                           Optional upToDate As Variant = "") As Currency
    Dim condition As String
    Dim result As Currency
    
    If product_id > 0 Then
        ' product_id takes precedence - aggregate across all product_spec_ids for this product_id
        condition = "product_id=" & product_id
    ElseIf product_spec_id > 0 Then
        ' Use specific product_spec_id
        condition = "product_spec_id=" & product_spec_id
    Else
        condition = "1=1"
    End If
    
    If IsDate(upToDate) Then
        condition = condition & " AND sale_date<=#" & Format(CDate(upToDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getItemReturnOutwardsValue] Using upToDate filter: <= " & Format(CDate(upToDate), "mm/dd/yyyy")
    ElseIf IsDate(startDate) And IsDate(endDate) Then
        condition = condition & " AND sale_date>=#" & Format(CDate(startDate), "mm/dd/yyyy") & "#"
        condition = condition & " AND sale_date<=#" & Format(CDate(endDate), "mm/dd/yyyy") & "#"
        Debug.Print "      [getItemReturnOutwardsValue] Using date range: " & Format(CDate(startDate), "mm/dd/yyyy") & " to " & Format(CDate(endDate), "mm/dd/yyyy")
    ElseIf month_id > 0 Then
        condition = condition & " AND month_id=" & month_id
        Debug.Print "      [getItemReturnOutwardsValue] Using month_id: " & month_id
    End If
    
    Debug.Print "      [getItemReturnOutwardsValue] Condition: " & condition
    result = Nz(DSum("amount", "qry_return_outwards", condition), 0)
    Debug.Print "      [getItemReturnOutwardsValue] Result: " & Format(result, "#,##0.00")
    
    getItemReturnOutwardsValue = result
End Function

' ==============================================================================
' SECTION 3: ITEM-LEVEL NET INVENTORY POSITION (SINGLE PRODUCT)
' Calculates actual quantity on hand considering ALL transactions
' ==============================================================================

' Net inventory quantity = Purchases + ReturnInwards - Sales - CreditSales - ReturnOutwards - OfficeUse - Drawings
Public Function getItemInventoryQty(product_spec_id As Long, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional endDate As Variant = "", _
                                    Optional upToDate As Variant = "") As Double
    Dim netQty As Double
    Dim qtyPurchased As Double
    Dim qtyReturnedInwards As Double
    Dim qtySold As Double
    Dim qtyCreditSales As Double
    Dim qtyReturnedOutwards As Double
    Dim qtyOfficeUse As Double
    Dim qtyDrawings As Double
    
    Debug.Print "    [getItemInventoryQty] product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "    [getItemInventoryQty] month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & endDate & ", upToDate: " & upToDate
    
    ' INFLOWS (goods coming IN to inventory)
    qtyPurchased = getItemQtyPurchased(product_spec_id, product_id, month_id, startDate, endDate, upToDate)
    Debug.Print "    [getItemInventoryQty] Purchased: " & qtyPurchased
    netQty = qtyPurchased
    
    qtyReturnedInwards = getItemQtyReturnedInwards(product_spec_id, product_id, month_id, startDate, endDate, upToDate)
    Debug.Print "    [getItemInventoryQty] Returned Inwards: " & qtyReturnedInwards
    netQty = netQty + qtyReturnedInwards
    
    ' OUTFLOWS (goods going OUT of inventory)
    qtySold = getItemQtySold(product_spec_id, product_id, month_id, startDate, endDate, upToDate)
    Debug.Print "    [getItemInventoryQty] Sold: " & qtySold
    netQty = netQty - qtySold
    
    qtyCreditSales = getItemQtyCreditSales(product_spec_id, product_id, month_id, startDate, endDate, upToDate)
    Debug.Print "    [getItemInventoryQty] Credit Sales: " & qtyCreditSales
    netQty = netQty - qtyCreditSales
    
    qtyReturnedOutwards = getItemQtyReturnedOutwards(product_spec_id, product_id, month_id, startDate, endDate, upToDate)
    Debug.Print "    [getItemInventoryQty] Returned Outwards: " & qtyReturnedOutwards
    netQty = netQty - qtyReturnedOutwards
    
    qtyOfficeUse = getItemQtyOfficeUse(product_spec_id, product_id, month_id, startDate, endDate, upToDate)
    Debug.Print "    [getItemInventoryQty] Office Use: " & qtyOfficeUse
    netQty = netQty - qtyOfficeUse
    
    qtyDrawings = getItemQtyDrawings(product_spec_id, product_id, month_id, startDate, endDate, upToDate)
    Debug.Print "    [getItemInventoryQty] Drawings: " & qtyDrawings
    netQty = netQty - qtyDrawings
    
    Debug.Print "    [getItemInventoryQty] Net Quantity = " & qtyPurchased & " + " & qtyReturnedInwards & " - " & qtySold & " - " & qtyCreditSales & " - " & qtyReturnedOutwards & " - " & qtyOfficeUse & " - " & qtyDrawings & " = " & netQty
    
    getItemInventoryQty = netQty
End Function

' ==============================================================================
' SECTION 4: ITEM-LEVEL WEIGHTED AVERAGE COST (SINGLE PRODUCT)
' Core costing method - ensures time-consistent historical valuations
' ==============================================================================

' Calculate weighted average cost per unit based on cumulative purchases up to date
' This is the FOUNDATION of all inventory valuation
' NOTE: When product_id is provided, calculates weighted average across all product_spec_ids for that product
Public Function getItemWeightedAvgCost(product_spec_id As Long, upToDate As Date, Optional product_id As Long = 0) As Double
    Dim totalCost As Currency
    Dim totalQty As Double
    Dim result As Double
    
    Debug.Print "  [getItemWeightedAvgCost] product_spec_id: " & product_spec_id & ", product_id: " & product_id & ", upToDate: " & Format(upToDate, "mm/dd/yyyy")
    
    totalCost = getItemPurchaseValue(product_spec_id, product_id, , , , upToDate)
    Debug.Print "  [getItemWeightedAvgCost] Total Purchase Cost up to " & Format(upToDate, "mm/dd/yyyy") & ": " & Format(totalCost, "#,##0.00")
    
    totalQty = getItemQtyPurchased(product_spec_id, product_id, , , , upToDate)
    Debug.Print "  [getItemWeightedAvgCost] Total Purchase Quantity up to " & Format(upToDate, "mm/dd/yyyy") & ": " & totalQty
    
    If totalQty = 0 Then
        Debug.Print "  [getItemWeightedAvgCost] TotalQty = 0, returning 0"
        getItemWeightedAvgCost = 0
    Else
        result = totalCost / totalQty
        Debug.Print "  [getItemWeightedAvgCost] Weighted Avg Cost (TotalCost/TotalQty): " & Format(result, "#,##0.00")
        getItemWeightedAvgCost = result
    End If
End Function

' ==============================================================================
' SECTION 5: ITEM-LEVEL OPENING & CLOSING STOCK (SINGLE PRODUCT)
' Balance sheet inventory positions at period boundaries
' ==============================================================================

' Opening stock QUANTITY at start of period
' EXPLANATION: Opening stock for a period = Closing stock from previous period
' Example: If period starts on 01/07/2025, opening stock = stock at end of 01/06/2025
'          (The stock quantity at 11:59:59 PM on 01/06/2025 is the same as 12:00:00 AM on 01/07/2025)
' Therefore: upToDate = periodStartDate - 1 to get all transactions up to end of previous day
Public Function getItemOpeningQty(product_spec_id As Long, periodStartDate As Date, Optional product_id As Long = 0) As Double
    Dim upToDate As Date
    Dim result As Double
    upToDate = periodStartDate - 1  ' Position at end of previous day (which equals start of period)
    Debug.Print "  [getItemOpeningQty] periodStartDate: " & Format(periodStartDate, "mm/dd/yyyy") & ", upToDate: " & Format(upToDate, "mm/dd/yyyy")
    Debug.Print "  [getItemOpeningQty] NOTE: upToDate is one day BEFORE periodStartDate because opening stock = closing stock of previous day"
    result = getItemInventoryQty(product_spec_id, product_id, , , , upToDate)
    Debug.Print "  [getItemOpeningQty] Opening Quantity (stock at end of " & Format(upToDate, "mm/dd/yyyy") & " = start of " & Format(periodStartDate, "mm/dd/yyyy") & "): " & result
    getItemOpeningQty = result
End Function

' Closing stock QUANTITY at end of period
Public Function getItemClosingQty(product_spec_id As Long, periodEndDate As Date, Optional product_id As Long = 0) As Double
    Dim result As Double
    Debug.Print "  [getItemClosingQty] periodEndDate: " & Format(periodEndDate, "mm/dd/yyyy")
    result = getItemInventoryQty(product_spec_id, product_id, , , , periodEndDate)
    Debug.Print "  [getItemClosingQty] Closing Quantity: " & result
    getItemClosingQty = result
End Function

' Opening stock VALUE at start of period (for balance sheet)
Public Function getItemOpeningStock(product_spec_id As Long, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional endDate As Variant = "") As Currency
    Dim periodStartDate As Date
    Dim priceDate As Date
    Dim openingQty As Double
    Dim avgCost As Double
    Dim result As Currency
    
    Debug.Print "=== getItemOpeningStock DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & endDate
    
    ' Determine period start date
    If IsDate(startDate) Then
        periodStartDate = CDate(startDate)
        Debug.Print "Using startDate: " & Format(periodStartDate, "mm/dd/yyyy")
    ElseIf month_id > 0 Then
        periodStartDate = DateSerial(Year(Date), month_id, 1)
        Debug.Print "Using month_id, periodStartDate: " & Format(periodStartDate, "mm/dd/yyyy")
    Else
        Debug.Print "No valid period start - returning 0"
        getItemOpeningStock = 0
        Exit Function
    End If
    
    openingQty = getItemOpeningQty(product_spec_id, periodStartDate, product_id)
    Debug.Print "Opening Quantity: " & openingQty
    
    If openingQty <= 0 Then
        Debug.Print "Opening Quantity <= 0 - returning 0"
        Debug.Print "=== END getItemOpeningStock DEBUG ==="
        getItemOpeningStock = 0
        Exit Function
    End If
    
    ' Value at weighted average cost as of day before period starts
    priceDate = periodStartDate - 1
    Debug.Print "Price Date (periodStartDate - 1): " & Format(priceDate, "mm/dd/yyyy")
    avgCost = getItemWeightedAvgCost(product_spec_id, priceDate, product_id)
    Debug.Print "Weighted Average Cost: " & Format(avgCost, "#,##0.00")
    
    result = CCur(openingQty * avgCost)
    Debug.Print "Opening Stock Value (Qty × AvgCost): " & Format(result, "#,##0.00")
    Debug.Print "=== END getItemOpeningStock DEBUG ==="
    
    getItemOpeningStock = result
End Function

' Closing stock VALUE at end of period (for balance sheet)
Public Function getItemClosingStock(product_spec_id As Long, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional endDate As Variant = "") As Currency
    Dim periodEndDate As Date
    Dim closingQty As Double
    Dim avgCost As Double
    Dim result As Currency
    
    Debug.Print "=== getItemClosingStock DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & endDate
    
    ' Determine period end date
    ' Handle endDate that might have trailing # or quotes
    Dim endDateStr As String
    If Not IsNull(endDate) And Not IsEmpty(endDate) And CStr(endDate) <> "" Then
        endDateStr = Trim(CStr(endDate))
        ' Remove trailing # and quotes
        endDateStr = Replace(endDateStr, "#", "")
        endDateStr = Replace(endDateStr, """", "")
        If IsDate(endDateStr) Then
            periodEndDate = CDate(endDateStr)
            Debug.Print "Using endDate (cleaned from '" & CStr(endDate) & "'): " & Format(periodEndDate, "mm/dd/yyyy")
        ElseIf month_id > 0 Then
            periodEndDate = DateSerial(Year(Date), month_id + 1, 0)
            Debug.Print "endDate invalid ('" & CStr(endDate) & "'), using month_id, periodEndDate: " & Format(periodEndDate, "mm/dd/yyyy")
        Else
            periodEndDate = Date
            Debug.Print "endDate invalid ('" & CStr(endDate) & "'), using current date as periodEndDate: " & Format(periodEndDate, "mm/dd/yyyy")
        End If
    ElseIf month_id > 0 Then
        periodEndDate = DateSerial(Year(Date), month_id + 1, 0)
        Debug.Print "Using month_id, periodEndDate: " & Format(periodEndDate, "mm/dd/yyyy")
    Else
        periodEndDate = Date
        Debug.Print "No endDate provided, using current date as periodEndDate: " & Format(periodEndDate, "mm/dd/yyyy")
    End If
    
    closingQty = getItemClosingQty(product_spec_id, periodEndDate, product_id)
    Debug.Print "Closing Quantity: " & closingQty
    
    If closingQty <= 0 Then
        Debug.Print "Closing Quantity <= 0 - returning 0"
        Debug.Print "=== END getItemClosingStock DEBUG ==="
        getItemClosingStock = 0
        Exit Function
    End If
    
    ' Value at weighted average cost as of period end
    Debug.Print "Price Date (periodEndDate): " & Format(periodEndDate, "mm/dd/yyyy")
    avgCost = getItemWeightedAvgCost(product_spec_id, periodEndDate, product_id)
    Debug.Print "Weighted Average Cost: " & Format(avgCost, "#,##0.00")
    
    result = CCur(closingQty * avgCost)
    Debug.Print "Closing Stock Value (Qty × AvgCost): " & Format(result, "#,##0.00")
    Debug.Print "=== END getItemClosingStock DEBUG ==="
    
    getItemClosingStock = result
End Function

' ==============================================================================
' SECTION 6: ITEM-LEVEL COGS (SINGLE PRODUCT)
' Cost of Goods Sold - the cost of ALL goods that left inventory
' ==============================================================================

' Calculate COGS for a single product in a period
' CORRECT FORMULA FOR WEIGHTED AVERAGE COSTING: COGS = Opening Stock + Net Purchases - Closing Stock
' This is the correct GAAP method because weighted average cost changes with each purchase
Public Function getItemCOGS(product_spec_id As Long, _
                            Optional product_id As Long = 0, _
                            Optional month_id As Long = 0, _
                            Optional startDate As Variant = "", _
                            Optional endDate As Variant = "") As Currency
    Dim openingStock As Currency
    Dim netPurchases As Currency
    Dim closingStock As Currency
    Dim cogs As Currency
    
    Debug.Print "=== getItemCOGS DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & endDate
    
    ' Get opening stock value
    openingStock = getItemOpeningStock(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "Opening Stock Value: " & Format(openingStock, "#,##0.00")
    
    ' Get net purchases for the period
    netPurchases = getItemNetPurchases(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "Net Purchases: " & Format(netPurchases, "#,##0.00")
    
    ' Get closing stock value
    closingStock = getItemClosingStock(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "Closing Stock Value: " & Format(closingStock, "#,##0.00")
    
    ' Calculate COGS using the inventory equation: Opening + Net Purchases - Closing = COGS
    cogs = openingStock + netPurchases - closingStock
    Debug.Print "COGS (Opening + NetPurchases - Closing): " & Format(cogs, "#,##0.00")
    Debug.Print "=== END getItemCOGS DEBUG ==="
    
    getItemCOGS = cogs
End Function

' ==============================================================================
' SECTION 7: ITEM-LEVEL ACCOUNTING CONCEPTS (SINGLE PRODUCT)
' Professional accounting calculations per GAAP
' ==============================================================================

' Net Purchases = Purchases - Purchase Returns
Public Function getItemNetPurchases(product_spec_id As Long, _
                                    Optional product_id As Long = 0, _
                                    Optional month_id As Long = 0, _
                                    Optional startDate As Variant = "", _
                                    Optional endDate As Variant = "") As Currency
    Dim purchases As Currency
    Dim returnOut As Currency
    Dim result As Currency
    
    Debug.Print "=== getItemNetPurchases DEBUG ==="
    Debug.Print "product_spec_id: " & product_spec_id & ", product_id: " & product_id
    Debug.Print "month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & endDate
    
    purchases = getItemPurchaseValue(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "Purchases Value: " & Format(purchases, "#,##0.00")
    
    returnOut = getItemReturnOutwardsValue(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "Return Outwards Value: " & Format(returnOut, "#,##0.00")
    
    result = purchases - returnOut
    Debug.Print "Net Purchases (Purchases - ReturnOutwards): " & Format(result, "#,##0.00")
    Debug.Print "=== END getItemNetPurchases DEBUG ==="
    
    getItemNetPurchases = result
End Function

' Net Sales = Cash Sales + Credit Sales - Sales Returns
' NOTE: Credit Sales are included in revenue (even though payment is deferred)
'       because revenue is recognized when goods are sold (accrual basis accounting)
'       Sales Returns reduce revenue because goods are returned by customers
Public Function getItemNetSales(product_spec_id As Long, _
                                Optional product_id As Long = 0, _
                                Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional endDate As Variant = "") As Currency
    Dim cashSales As Currency
    Dim creditSales As Currency
    Dim returnIn As Currency
    
    cashSales = getItemSalesRevenue(product_spec_id, product_id, month_id, startDate, endDate)
    creditSales = getItemCreditSalesRevenue(product_spec_id, product_id, month_id, startDate, endDate)
    returnIn = getItemReturnInwardsValue(product_spec_id, product_id, month_id, startDate, endDate)
    
    getItemNetSales = cashSales + creditSales - returnIn
End Function

' Gross Profit = Net Sales - COGS
Public Function getItemGrossProfit(product_spec_id As Long, _
                                   Optional product_id As Long = 0, _
                                   Optional month_id As Long = 0, _
                                   Optional startDate As Variant = "", _
                                   Optional endDate As Variant = "") As Currency
    Dim netSales As Currency
    Dim cogs As Currency
    
    netSales = getItemNetSales(product_spec_id, product_id, month_id, startDate, endDate)
    cogs = getItemCOGS(product_spec_id, product_id, month_id, startDate, endDate)
    
    getItemGrossProfit = netSales - cogs
End Function

' Gross Profit Margin (percentage)
Public Function getItemGrossMargin(product_spec_id As Long, _
                                   Optional product_id As Long = 0, _
                                   Optional month_id As Long = 0, _
                                   Optional startDate As Variant = "", _
                                   Optional endDate As Variant = "") As Double
    Dim netSales As Currency
    Dim grossProfit As Currency
    
    netSales = getItemNetSales(product_spec_id, product_id, month_id, startDate, endDate)
    
    If netSales = 0 Then
        getItemGrossMargin = 0
    Else
        grossProfit = getItemGrossProfit(product_spec_id, product_id, month_id, startDate, endDate)
        getItemGrossMargin = Round((grossProfit / netSales) * 100, 2)
    End If
End Function

' Cost of Goods Available for Sale = Opening Stock + Net Purchases
Public Function getItemCOGAS(product_spec_id As Long, _
                             Optional product_id As Long = 0, _
                             Optional month_id As Long = 0, _
                             Optional startDate As Variant = "", _
                             Optional endDate As Variant = "") As Currency
    Dim openingStock As Currency
    Dim netPurchases As Currency
    
    openingStock = getItemOpeningStock(product_spec_id, product_id, month_id, startDate, endDate)
    netPurchases = getItemNetPurchases(product_spec_id, product_id, month_id, startDate, endDate)
    
    getItemCOGAS = openingStock + netPurchases
End Function

' ==============================================================================
' SECTION 8: AGGREGATE FUNCTIONS (ALL PRODUCTS)
' These iterate over all products and sum results
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

' Total Opening Stock - all products (Balance Sheet: Current Assets)
Public Function getOpeningStock(Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional endDate As Variant = "") As Currency
    Dim products As Collection
    Dim productID As Variant
    Dim total As Currency
    
    total = 0
    Set products = getAllProductSpecIDs()
    
    For Each productID In products
        total = total + getItemOpeningStock(CLng(productID), 0, month_id, startDate, endDate)
    Next productID
    
    getOpeningStock = total
End Function

' Total Closing Stock - all products (Balance Sheet: Current Assets)
Public Function getClosingStock(Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional endDate As Variant = "") As Currency
    Dim products As Collection
    Dim productID As Variant
    Dim total As Currency
    
    total = 0
    Set products = getAllProductSpecIDs()
    
    For Each productID In products
        total = total + getItemClosingStock(CLng(productID), 0, month_id, startDate, endDate)
    Next productID
    
    getClosingStock = total
End Function

' Total Purchases - all products
Public Function getPurchases(Optional month_id As Long = 0, _
                             Optional startDate As Variant = "", _
                             Optional endDate As Variant = "") As Currency
    Dim products As Collection
    Dim productID As Variant
    Dim total As Currency
    
    total = 0
    Set products = getAllProductSpecIDs()
    
    For Each productID In products
        total = total + getItemPurchaseValue(CLng(productID), 0, month_id, startDate, endDate)
    Next productID
    
    getPurchases = total
End Function

' Total Net Purchases - all products (Income Statement)
Public Function getNetPurchases(Optional month_id As Long = 0, _
                                Optional startDate As Variant = "", _
                                Optional endDate As Variant = "") As Currency
    Dim products As Collection
    Dim productID As Variant
    Dim total As Currency
    
    total = 0
    Set products = getAllProductSpecIDs()
    
    For Each productID In products
        total = total + getItemNetPurchases(CLng(productID), 0, month_id, startDate, endDate)
    Next productID
    
    getNetPurchases = total
End Function

' Total Net Sales - all products (Income Statement: Revenue)
Public Function getNetSales(Optional month_id As Long = 0, _
                            Optional startDate As Variant = "", _
                            Optional endDate As Variant = "") As Currency
    Dim products As Collection
    Dim productID As Variant
    Dim total As Currency
    
    total = 0
    Set products = getAllProductSpecIDs()
    
    For Each productID In products
        total = total + getItemNetSales(CLng(productID), 0, month_id, startDate, endDate)
    Next productID
    
    getNetSales = total
End Function

' Total COGS - all products (Income Statement)
Public Function getCOGS(Optional month_id As Long = 0, _
                        Optional startDate As Variant = "", _
                        Optional endDate As Variant = "") As Currency
    Dim products As Collection
    Dim productID As Variant
    Dim total As Currency
    
    total = 0
    Set products = getAllProductSpecIDs()
    
    For Each productID In products
        total = total + getItemCOGS(CLng(productID), 0, month_id, startDate, endDate)
    Next productID
    
    getCOGS = total
End Function

' Total Gross Profit - all products (Income Statement)
Public Function getGrossProfit(Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional endDate As Variant = "") As Currency
    Dim netSales As Currency
    Dim cogs As Currency
    
    netSales = getNetSales(month_id, startDate, endDate)
    cogs = getCOGS(month_id, startDate, endDate)
    
    getGrossProfit = netSales - cogs
End Function

' Overall Gross Margin (percentage)
Public Function getGrossMargin(Optional month_id As Long = 0, _
                               Optional startDate As Variant = "", _
                               Optional endDate As Variant = "") As Double
    Dim netSales As Currency
    Dim grossProfit As Currency
    
    netSales = getNetSales(month_id, startDate, endDate)
    
    If netSales = 0 Then
        getGrossMargin = 0
    Else
        grossProfit = getGrossProfit(month_id, startDate, endDate)
        getGrossMargin = Round((grossProfit / netSales) * 100, 2)
    End If
End Function

' Total COGAS - all products
Public Function getCOGAS(Optional month_id As Long = 0, _
                         Optional startDate As Variant = "", _
                         Optional endDate As Variant = "") As Currency
    Dim products As Collection
    Dim productID As Variant
    Dim total As Currency
    
    total = 0
    Set products = getAllProductSpecIDs()
    
    For Each productID In products
        total = total + getItemCOGAS(CLng(productID), 0, month_id, startDate, endDate)
    Next productID
    
    getCOGAS = total
End Function

' ==============================================================================
' SECTION 9: VALIDATION & REPORTING FUNCTIONS
' ==============================================================================

' Validate inventory equation: Opening + Net Purchases = COGS + Closing
' NOTE: Since COGS is now calculated as Opening + Net Purchases - Closing,
' this validation will always pass mathematically. It's kept for verification.
Public Function ValidateItemInventoryEquation(product_spec_id As Long, _
                                              Optional product_id As Long = 0, _
                                              Optional month_id As Long = 0, _
                                              Optional startDate As Variant = "", _
                                              Optional endDate As Variant = "") As Boolean
    Dim openingStock As Currency
    Dim netPurchases As Currency
    Dim cogs As Currency
    Dim closingStock As Currency
    Dim leftSide As Currency
    Dim rightSide As Currency
    Dim difference As Currency
    
    Debug.Print vbCrLf & "========================================"
    Debug.Print "INVENTORY EQUATION VALIDATION START"
    Debug.Print "========================================"
    If product_id > 0 Then
        Debug.Print "Product ID: " & product_id & " (Product Spec ID: " & product_spec_id & ")"
    Else
        Debug.Print "Product Spec ID: " & product_spec_id
    End If
    Debug.Print "month_id: " & month_id & ", startDate: " & startDate & ", endDate: " & endDate
    Debug.Print "----------------------------------------"
    
    openingStock = getItemOpeningStock(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "----------------------------------------"
    
    netPurchases = getItemNetPurchases(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "----------------------------------------"
    
    cogs = getItemCOGS(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "----------------------------------------"
    
    closingStock = getItemClosingStock(product_spec_id, product_id, month_id, startDate, endDate)
    Debug.Print "----------------------------------------"
    
    leftSide = openingStock + netPurchases
    rightSide = cogs + closingStock
    difference = Abs(leftSide - rightSide)
    
    Debug.Print "----------------------------------------"
    Debug.Print "VALIDATION SUMMARY:"
    Debug.Print "Opening Stock: " & Format(openingStock, "#,##0.00")
    Debug.Print "Net Purchases: " & Format(netPurchases, "#,##0.00")
    Debug.Print "COGS: " & Format(cogs, "#,##0.00")
    Debug.Print "Closing Stock: " & Format(closingStock, "#,##0.00")
    Debug.Print "Left Side (Opening + Net Purchases): " & Format(leftSide, "#,##0.00")
    Debug.Print "Right Side (COGS + Closing): " & Format(rightSide, "#,##0.00")
    Debug.Print "Difference: " & Format(difference, "#,##0.00")
    
    ' Allow 0.01 rounding tolerance
    ValidateItemInventoryEquation = (difference < 0.01)
    Debug.Print "Status: " & IIf(difference < 0.01, "PASSED", "FAILED")
    Debug.Print "========================================"
    Debug.Print "INVENTORY EQUATION VALIDATION END"
    Debug.Print "========================================" & vbCrLf
    
End Function

