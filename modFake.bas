Attribute VB_Name = "modFake"
Option Compare Database
Option Explicit

' ========================================
' MODULE: modFakeDataGenerator
' Purpose: Generate fake sales and purchase data for testing
' ========================================

' ========================================
' SALES DATA GENERATION
' ========================================

Public Sub GenerateFakeSalesForDate(Optional sale_date As Variant, Optional month_id As Long = 0)
    ' This function generates random sales for a given date or entire month
    ' If month_id is provided (1-12), generates sales for all days in that month
    ' If sale_date is provided, generates sales for that specific date
    ' If neither provided, generates sales for today
    ' Prioritizes selling cheaper items more frequently than expensive items
    
    Dim db As DAO.Database
    Dim rsProducts As DAO.Recordset
    Dim sqlQuery As String
    Dim totalProducts As Long
    Dim numSalesToGenerate As Long
    Dim i As Long
    Dim randomIndex As Long
    Dim product_spec_id As Long
    Dim current_stock As Long
    Dim unit_price As Currency
    Dim quantity As Long
    Dim saleResult As Long
    Dim currentDate As Date
    Dim startDate As Date
    Dim EndDate As Date
    Dim daysInMonth As Long
    Dim dayCounter As Long
    Dim totalSalesCreated As Long
    Dim priceCategory As String
    
    On Error GoTo ErrorHandler
    
    ' Determine date range based on parameters
    If month_id >= 1 And month_id <= 12 Then
        startDate = DateSerial(Year(Date), month_id, 1)
        EndDate = DateSerial(Year(Date), month_id + 1, 0)
        daysInMonth = Day(EndDate)
        Debug.Print "Generating sales for entire month: " & Format(startDate, "mmmm yyyy")
    ElseIf Not IsNull(sale_date) And Not IsEmpty(sale_date) Then
        startDate = CDate(sale_date)
        EndDate = startDate
        daysInMonth = 1
    Else
        startDate = Date
        EndDate = startDate
        daysInMonth = 1
    End If
    
    totalSalesCreated = 0
    Set db = CurrentDb()
    
    ' SQL to get products with price-based ordering (cheaper items first)
    sqlQuery = "SELECT DISTINCT q.product_spec_id, q.current_stock, " & _
               "p.suggested_selling_price, " & _
               "IIF(p.suggested_selling_price < 5000, 'cheap', " & _
               "IIF(p.suggested_selling_price < 10000, 'medium', 'expensive')) AS price_category " & _
               "FROM qry_all_products AS q " & _
               "INNER JOIN tbl_purchase_details AS p ON q.product_spec_id = p.product_spec_id " & _
               "WHERE q.current_stock > 0 " & _
               "ORDER BY p.suggested_selling_price ASC"
    
    Set rsProducts = db.OpenRecordset(sqlQuery, dbOpenSnapshot)
    
    If rsProducts.EOF Then
        MsgBox "No products available with stock and purchase details.", vbExclamation
        GoTo Cleanup
    End If
    
    rsProducts.MoveLast
    totalProducts = rsProducts.RecordCount
    rsProducts.MoveFirst
    
    ' Loop through each day in the date range
    For dayCounter = 0 To daysInMonth - 1
        currentDate = DateAdd("d", dayCounter, startDate)
        
        Randomize Timer + dayCounter
        numSalesToGenerate = Int((totalProducts * 0.4 - 3 + 1) * Rnd + 3)
        If numSalesToGenerate < 3 Then numSalesToGenerate = 3
        
        Debug.Print "Generating " & numSalesToGenerate & " sales for " & Format(currentDate, "yyyy-mm-dd")
        
        ' Generate sales with price-weighted selection
        For i = 1 To numSalesToGenerate
            ' Weighted selection: 70% cheap, 20% medium, 10% expensive
            Dim randWeight As Double
            randWeight = Rnd
            
            rsProducts.Requery
            If rsProducts.EOF Then Exit For
            
            If randWeight < 0.7 Then
                priceCategory = "cheap"
            ElseIf randWeight < 0.9 Then
                priceCategory = "medium"
            Else
                priceCategory = "expensive"
            End If
            
            ' Find products in the selected price category
            rsProducts.MoveFirst
            Dim matchingProducts As Long
            Dim startPosition As Long
            matchingProducts = 0
            startPosition = 0
            
            Do While Not rsProducts.EOF
                If rsProducts!price_category = priceCategory Then
                    If matchingProducts = 0 Then
                        startPosition = rsProducts.AbsolutePosition
                    End If
                    matchingProducts = matchingProducts + 1
                End If
                rsProducts.MoveNext
            Loop
            
            ' Select product
            If matchingProducts = 0 Then
                randomIndex = Int(totalProducts * Rnd + 1)
                If randomIndex < 1 Then randomIndex = 1
                rsProducts.MoveFirst
                rsProducts.Move randomIndex - 1
            Else
                randomIndex = Int(matchingProducts * Rnd)
                rsProducts.MoveFirst
                rsProducts.Move startPosition + randomIndex
            End If
            
            ' Get product details
            product_spec_id = rsProducts!product_spec_id
            current_stock = rsProducts!current_stock
            unit_price = rsProducts!suggested_selling_price
            
            ' Generate quantity based on price
            If unit_price < 5000 Then
                quantity = Int((current_stock * 0.3 - current_stock * 0.1 + 1) * Rnd + current_stock * 0.1)
            ElseIf unit_price < 10000 Then
                quantity = Int((current_stock * 0.2 - current_stock * 0.05 + 1) * Rnd + current_stock * 0.05)
            Else
                quantity = Int((current_stock * 0.1 - current_stock * 0.01 + 1) * Rnd + current_stock * 0.01)
            End If
            
            If quantity < 1 Then quantity = 1
            If quantity > current_stock Then quantity = current_stock
            
            ' Create the sale
            saleResult = CreateSale( _
                product_spec_id:=product_spec_id, _
                quantity:=quantity, _
                unit_price:=unit_price, _
                sale_date:=currentDate _
            )
            
            If saleResult > 0 Then
                totalSalesCreated = totalSalesCreated + 1
                Debug.Print "Created sale ID: " & saleResult & " - Product: " & product_spec_id & _
                           " (" & priceCategory & "), Qty: " & quantity & ", Price: " & Format(unit_price, "Currency")
            End If
            
            rsProducts.Requery
            If rsProducts.EOF Then Exit For
            
        Next i
    Next dayCounter
    
    If daysInMonth > 1 Then
        Debug.Print "Successfully generated " & totalSalesCreated & " fake sales for " & _
               Format(startDate, "mmmm yyyy") & " (" & daysInMonth & " days)", vbInformation
    Else
        Debug.Print "Successfully generated " & totalSalesCreated & " fake sales for " & _
               Format(startDate, "yyyy-mm-dd"), vbInformation
    End If
    
Cleanup:
    If Not rsProducts Is Nothing Then rsProducts.Close
    Set rsProducts = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    'MsgBox "Error generating fake sales: " & Err.description, vbCritical
    Resume Cleanup
End Sub

' ========================================
' PURCHASE DATA GENERATION
' ========================================

Public Sub GenerateFakePurchasesForDate(Optional purchase_date As Variant, Optional month_id As Long = 0)
    ' Generates random purchases for a given date or entire month
    ' Prioritizes products with low stock and cheaper items over expensive ones
    
    Dim db As DAO.Database
    Dim rsSuppliers As DAO.Recordset
    Dim rsProducts As DAO.Recordset
    Dim sqlQuery As String
    Dim totalSuppliers As Long
    Dim totalProducts As Long
    Dim numPurchasesToGenerate As Long
    Dim i As Long
    Dim randomIndex As Long
    Dim supplier_id As Long
    Dim product_spec_id As Long
    Dim current_stock As Long
    Dim unit_cost As Currency
    Dim suggested_selling_price As Currency
    Dim quantity As Long
    Dim purchase_id As Long
    Dim currentDate As Date
    Dim startDate As Date
    Dim EndDate As Date
    Dim daysInMonth As Long
    Dim dayCounter As Long
    Dim totalPurchasesCreated As Long
    Dim priceCategory As String
    
    On Error GoTo ErrorHandler
    
    ' Determine date range
    If month_id >= 1 And month_id <= 12 Then
        startDate = DateSerial(Year(Date), month_id, 1)
        EndDate = DateSerial(Year(Date), month_id + 1, 0)
        daysInMonth = Day(EndDate)
        Debug.Print "Generating purchases for entire month: " & Format(startDate, "mmmm yyyy")
    ElseIf Not IsNull(purchase_date) And Not IsEmpty(purchase_date) Then
        startDate = CDate(purchase_date)
        EndDate = startDate
        daysInMonth = 1
    Else
        startDate = Date
        EndDate = startDate
        daysInMonth = 1
    End If
    
    totalPurchasesCreated = 0
    Set db = CurrentDb()
    
    ' Get suppliers
    Set rsSuppliers = db.OpenRecordset("SELECT supplier_id FROM tbl_suppliers", dbOpenSnapshot)
    
    If rsSuppliers.EOF Then
        MsgBox "No suppliers available in tbl_suppliers.", vbExclamation
        GoTo Cleanup
    End If
    
    rsSuppliers.MoveLast
    totalSuppliers = rsSuppliers.RecordCount
    rsSuppliers.MoveFirst
    
    ' Get products with price and stock categorization
    sqlQuery = "SELECT DISTINCT pd.product_spec_id, " & _
               "q.current_stock, " & _
               "pd.unit_cost, " & _
               "pd.suggested_selling_price, " & _
               "IIF(pd.unit_cost < 5000, 'cheap', " & _
               "IIF(pd.unit_cost < 10000, 'medium', 'expensive')) AS price_category, " & _
               "IIF(q.current_stock < 50, 'critical', " & _
               "IIF(q.current_stock < 100, 'low', " & _
               "IIF(q.current_stock < 200, 'medium', 'high'))) AS stock_level " & _
               "FROM tbl_purchase_details AS pd " & _
               "INNER JOIN qry_all_products AS q ON pd.product_spec_id = q.product_spec_id " & _
               "WHERE pd.unit_cost > 0 AND pd.suggested_selling_price > 0 " & _
               "ORDER BY q.current_stock ASC, pd.unit_cost ASC"
    
    Set rsProducts = db.OpenRecordset(sqlQuery, dbOpenSnapshot)
    
    If rsProducts.EOF Then
        MsgBox "No products available with existing purchase details.", vbExclamation
        GoTo Cleanup
    End If
    
    rsProducts.MoveLast
    totalProducts = rsProducts.RecordCount
    rsProducts.MoveFirst
    
    ' Loop through each day
    For dayCounter = 0 To daysInMonth - 1
        currentDate = DateAdd("d", dayCounter, startDate)
        
        Randomize Timer + dayCounter
        numPurchasesToGenerate = Int((5 - 1 + 1) * Rnd + 1)
        
        Debug.Print "Generating " & numPurchasesToGenerate & " purchase orders for " & Format(currentDate, "yyyy-mm-dd")
        
        ' Generate purchase orders
        For i = 1 To numPurchasesToGenerate
            ' Select random supplier
            randomIndex = Int(totalSuppliers * Rnd + 1)
            rsSuppliers.MoveFirst
            rsSuppliers.Move randomIndex - 1
            supplier_id = rsSuppliers!supplier_id
            
            ' Create purchase order
            purchase_id = CreatePurchaseOrder(supplier_id, currentDate)
            
            If purchase_id > 0 Then
                ' Add 2-6 products to this order
                Dim numItems As Long
                numItems = Int((6 - 2 + 1) * Rnd + 2)
                
                Dim j As Long
                For j = 1 To numItems
                    rsProducts.Requery
                    If rsProducts.EOF Then Exit For
                    
                    ' Weighted selection: 50% low stock+cheap/medium, 25% any stock+cheap,
                    ' 15% low stock+any price, 10% random
                    Dim randWeight As Double
                    randWeight = Rnd
                    
                    rsProducts.MoveFirst
                    Dim matchCount As Long
                    Dim matchStart As Long
                    matchCount = 0
                    matchStart = 0
                    
                    If randWeight < 0.5 Then
                        ' Low/Critical stock + Cheap/Medium price
                        Do While Not rsProducts.EOF
                            If (rsProducts!stock_level = "critical" Or rsProducts!stock_level = "low") And _
                               (rsProducts!price_category = "cheap" Or rsProducts!price_category = "medium") Then
                                If matchCount = 0 Then matchStart = rsProducts.AbsolutePosition
                                matchCount = matchCount + 1
                            End If
                            rsProducts.MoveNext
                        Loop
                    ElseIf randWeight < 0.75 Then
                        ' Any stock + Cheap items only
                        Do While Not rsProducts.EOF
                            If rsProducts!price_category = "cheap" Then
                                If matchCount = 0 Then matchStart = rsProducts.AbsolutePosition
                                matchCount = matchCount + 1
                            End If
                            rsProducts.MoveNext
                        Loop
                    ElseIf randWeight < 0.9 Then
                        ' Low/Critical stock + Any price
                        Do While Not rsProducts.EOF
                            If rsProducts!stock_level = "critical" Or rsProducts!stock_level = "low" Then
                                If matchCount = 0 Then matchStart = rsProducts.AbsolutePosition
                                matchCount = matchCount + 1
                            End If
                            rsProducts.MoveNext
                        Loop
                    End If
                    
                    ' Select product
                    If matchCount = 0 Then
                        randomIndex = Int(totalProducts * Rnd + 1)
                        If randomIndex < 1 Then randomIndex = 1
                        rsProducts.MoveFirst
                        rsProducts.Move randomIndex - 1
                    Else
                        randomIndex = Int(matchCount * Rnd)
                        rsProducts.MoveFirst
                        rsProducts.Move matchStart + randomIndex
                    End If
                    
                    ' Get product details
                    product_spec_id = rsProducts!product_spec_id
                    current_stock = rsProducts!current_stock
                    unit_cost = rsProducts!unit_cost
                    suggested_selling_price = rsProducts!suggested_selling_price
                    priceCategory = rsProducts!price_category
                    
                    ' Calculate quantity based on stock and price
                    Dim baseQty As Long
                    Dim maxQty As Long
                    
                    ' Base on stock level
                    If current_stock < 50 Then
                        baseQty = 50: maxQty = 150
                    ElseIf current_stock < 100 Then
                        baseQty = 30: maxQty = 80
                    ElseIf current_stock < 200 Then
                        baseQty = 15: maxQty = 40
                    Else
                        baseQty = 5: maxQty = 20
                    End If
                    
                    ' Adjust by price
                    If unit_cost >= 10000 Then
                        baseQty = Int(baseQty * 0.3)
                        maxQty = Int(maxQty * 0.3)
                        If baseQty < 1 Then baseQty = 1
                        If maxQty < 2 Then maxQty = 2
                    ElseIf unit_cost >= 5000 Then
                        baseQty = Int(baseQty * 0.6)
                        maxQty = Int(maxQty * 0.6)
                        If baseQty < 2 Then baseQty = 2
                        If maxQty < 5 Then maxQty = 5
                    End If
                    
                    quantity = Int((maxQty - baseQty + 1) * Rnd + baseQty)
                    If quantity < 1 Then quantity = 1
                    
                    ' Add purchase detail
                    If AddPurchaseDetail(purchase_id, product_spec_id, quantity, unit_cost, suggested_selling_price) Then
                        Debug.Print "  Added to purchase " & purchase_id & ": Product " & product_spec_id & _
                                   " (" & priceCategory & "), Stock: " & current_stock & _
                                   ", Qty: " & quantity & ", Cost: " & Format(unit_cost, "Currency")
                    End If
                Next j
                
                totalPurchasesCreated = totalPurchasesCreated + 1
                Debug.Print "Created purchase order ID: " & purchase_id & " from Supplier: " & supplier_id
            End If
        Next i
    Next dayCounter
    
    UpdateProductsStock
    If daysInMonth > 1 Then
        Debug.Print "Successfully generated " & totalPurchasesCreated & " fake purchase orders for " & _
               Format(startDate, "mmmm yyyy") & " (" & daysInMonth & " days)", vbInformation
    Else
        Debug.Print "Successfully generated " & totalPurchasesCreated & " fake purchase orders for " & _
               Format(startDate, "yyyy-mm-dd"), vbInformation
    End If
    
Cleanup:
    If Not rsProducts Is Nothing Then rsProducts.Close
    If Not rsSuppliers Is Nothing Then rsSuppliers.Close
    Set rsProducts = Nothing
    Set rsSuppliers = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating fake purchases: " & Err.description, vbCritical
    Resume Cleanup
End Sub

' ========================================
' HELPER FUNCTIONS
' ========================================

Private Function CreatePurchaseOrder(supplier_id As Long, purchase_date As Date) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tbl_purchases", dbOpenDynaset)
    
    rs.AddNew
    rs!supplier_id = supplier_id
    rs!purchase_date = purchase_date
    rs.upDate
    
    rs.Bookmark = rs.LastModified
    CreatePurchaseOrder = rs!purchase_id
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    CreatePurchaseOrder = 0
End Function

Private Function AddPurchaseDetail(purchase_id As Long, product_spec_id As Long, _
                                   quantity As Long, unit_cost As Currency, _
                                   suggested_selling_price As Currency) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tbl_purchase_details", dbOpenDynaset)
    
    rs.AddNew
    rs!purchase_id = purchase_id
    rs!product_spec_id = product_spec_id
    rs!quantity = quantity
    rs!unit_cost = unit_cost
    rs!suggested_selling_price = suggested_selling_price
    rs.upDate
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    AddPurchaseDetail = True
    Exit Function
    
ErrorHandler:
    AddPurchaseDetail = False
End Function

Public Sub dailySales()
    Dim i As Long
    Dim j As Integer
    
    For i = 7 To 12
        For j = 1 To 31
            GenerateFakeSalesForDate DateSerial(2025, i, j)
        Next j
        GenerateFakePurchasesForDate DateSerial(2025, i, Int(Rnd * 31))
    Next i
End Sub
























' Generate realistic debts with immediate payment processing
' Each debt is fully processed before creating the next one
Public Sub GenerateFakeDebtsAndReturns(startDate As Date, EndDate As Date)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim currentMonthStart As Date
    Dim currentMonthEnd As Date
    Dim totalDebtsCreated As Long
    Dim totalReturnsCreated As Long
    Dim totalFullyPaid As Long
    Dim totalPartiallyPaid As Long
    Dim totalUnpaid As Long
    
    Set db = CurrentDb
    totalDebtsCreated = 0
    totalReturnsCreated = 0
    totalFullyPaid = 0
    totalPartiallyPaid = 0
    totalUnpaid = 0
    
    ' Initialize random number generator
    Randomize Timer
    
    currentMonthStart = DateSerial(Year(startDate), Month(startDate), 1)
    
    ' Begin transaction
    DBEngine.BeginTrans
    
    Debug.Print "========================================="
    Debug.Print "STARTING DEBT GENERATION"
    Debug.Print "Date Range: " & Format(startDate, "dd/mm/yyyy") & " to " & Format(EndDate, "dd/mm/yyyy")
    Debug.Print "========================================="
    
    ' Process month by month
    Do While currentMonthStart <= EndDate
        ' Get last day of current month
        currentMonthEnd = DateSerial(Year(currentMonthStart), Month(currentMonthStart) + 1, 0)
        
        ' Don't go beyond the end date
        If currentMonthEnd > EndDate Then
            currentMonthEnd = EndDate
        End If
        
        ' Generate debts for this month - each debt fully processed before next
        Debug.Print ""
        Debug.Print "???????????????????????????????????????"
        Debug.Print "MONTH: " & Format(currentMonthStart, "mmmm yyyy")
        Debug.Print "???????????????????????????????????????"
        
        Call ProcessMonthDebts(currentMonthStart, currentMonthEnd, EndDate, _
                              totalDebtsCreated, totalReturnsCreated, _
                              totalFullyPaid, totalPartiallyPaid, totalUnpaid)
        
        ' Move to next month
        currentMonthStart = DateAdd("m", 1, currentMonthStart)
    Loop
    
    ' Commit transaction
    DBEngine.CommitTrans
    
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "GENERATION COMPLETE!"
    Debug.Print "========================================="
    Debug.Print "Total Debts Created: " & totalDebtsCreated
    Debug.Print "Total Returns Created: " & totalReturnsCreated
    Debug.Print "Fully Paid Debts: " & totalFullyPaid & " (" & Format(totalFullyPaid / totalDebtsCreated, "0.0%") & ")"
    Debug.Print "Partially Paid Debts: " & totalPartiallyPaid & " (" & Format(totalPartiallyPaid / totalDebtsCreated, "0.0%") & ")"
    Debug.Print "Unpaid Debts: " & totalUnpaid & " (" & Format(totalUnpaid / totalDebtsCreated, "0.0%") & ")"
    Debug.Print "========================================="
    
    Set db = Nothing
    
    MsgBox "Debt generation complete!" & vbCrLf & vbCrLf & _
           "Date range: " & Format(startDate, "dd/mm/yyyy") & " to " & Format(EndDate, "dd/mm/yyyy") & vbCrLf & _
           "Total Debts: " & totalDebtsCreated & vbCrLf & _
           "Total Returns: " & totalReturnsCreated & vbCrLf & vbCrLf & _
           "Fully Paid: " & totalFullyPaid & " (" & Format(totalFullyPaid / totalDebtsCreated, "0.0%") & ")" & vbCrLf & _
           "Partially Paid: " & totalPartiallyPaid & " (" & Format(totalPartiallyPaid / totalDebtsCreated, "0.0%") & ")" & vbCrLf & _
           "Unpaid: " & totalUnpaid & " (" & Format(totalUnpaid / totalDebtsCreated, "0.0%") & ")", _
           vbInformation
    Exit Sub
    
ErrorHandler:
    ' Rollback transaction on error
    On Error Resume Next
    DBEngine.Rollback
    On Error GoTo 0
    
    MsgBox "Error: " & Err.Number & " - " & Err.description & vbCrLf & _
           "At line: " & Erl, vbCritical
    Set db = Nothing
End Sub

' Process all debts for a specific month
Private Sub ProcessMonthDebts(monthStart As Date, monthEnd As Date, finalEndDate As Date, _
                              ByRef totalDebts As Long, ByRef totalReturns As Long, _
                              ByRef totalFullyPaid As Long, ByRef totalPartiallyPaid As Long, _
                              ByRef totalUnpaid As Long)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsDebtors As DAO.Recordset
    Dim debtorIds() As Long
    Dim numDebtorsToUse As Integer
    Dim totalDebtorsCount As Integer
    Dim selectedDebtors As Integer
    Dim i As Integer
    
    Set db = CurrentDb
    
    ' Get all debtors
    Set rsDebtors = db.OpenRecordset("SELECT debtor_id FROM tbl_debtors ORDER BY debtor_id", dbOpenSnapshot)
    
    If rsDebtors.EOF Then
        Debug.Print "  No debtors found!"
        rsDebtors.Close
        Set rsDebtors = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    ' Count total debtors
    rsDebtors.MoveLast
    totalDebtorsCount = rsDebtors.RecordCount
    rsDebtors.MoveFirst
    
    Debug.Print "Total debtors available: " & totalDebtorsCount
    
    ' Randomly select 50-100% of debtors for this month
    numDebtorsToUse = Int((totalDebtorsCount - Int(totalDebtorsCount / 2) + 1) * Rnd + Int(totalDebtorsCount / 2))
    If numDebtorsToUse < 1 Then numDebtorsToUse = 1
    
    Debug.Print "Will create debts for " & numDebtorsToUse & " debtors this month"
    Debug.Print ""
    
    ' Randomly select which debtors
    ReDim debtorIds(1 To numDebtorsToUse)
    selectedDebtors = 0
    
    Do While Not rsDebtors.EOF And selectedDebtors < numDebtorsToUse
        ' Random chance to include this debtor
        If Rnd < (CDbl(numDebtorsToUse - selectedDebtors) / CDbl(totalDebtorsCount - rsDebtors.AbsolutePosition)) Then
            selectedDebtors = selectedDebtors + 1
            debtorIds(selectedDebtors) = rsDebtors!debtor_id
        End If
        rsDebtors.MoveNext
    Loop
    
    rsDebtors.Close
    Set rsDebtors = Nothing
    
    ' Process each debtor - create debts and immediately process payments
    For i = 1 To selectedDebtors
        Call ProcessDebtorDebts(debtorIds(i), monthStart, monthEnd, finalEndDate, _
                               totalDebts, totalReturns, _
                               totalFullyPaid, totalPartiallyPaid, totalUnpaid)
    Next i
    
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ProcessMonthDebts: " & Err.description
    If Not rsDebtors Is Nothing Then
        If Not rsDebtors.EOF Then rsDebtors.Close
    End If
    Set rsDebtors = Nothing
    Set db = Nothing
End Sub

' Process debts for a single debtor in a month
Private Sub ProcessDebtorDebts(debtorId As Long, monthStart As Date, monthEnd As Date, _
                               finalEndDate As Date, ByRef totalDebts As Long, _
                               ByRef totalReturns As Long, ByRef totalFullyPaid As Long, _
                               ByRef totalPartiallyPaid As Long, ByRef totalUnpaid As Long)
    On Error GoTo ErrorHandler
    
    Dim numDebtsForDebtor As Integer
    Dim j As Integer
    Dim saleDate As Date
    Dim productSpecId As Long
    Dim debtId As Long
    Dim totalAmount As Currency
    Dim expectedPaymentDate As Date
    Dim paymentScenario As Integer
    Dim returnsCreated As Long
    Dim totalPaid As Currency
    Dim paymentStatus As String
    
    ' Random number of debts for this debtor this month (1-3)
    numDebtsForDebtor = Int((3 - 1 + 1) * Rnd + 1)
    
    Debug.Print "+- Debtor ID: " & debtorId & " (" & numDebtsForDebtor & " debt(s))"
    
    ' Create each debt and immediately process its payments
    For j = 1 To numDebtsForDebtor
        ' Random sale date within this month
        saleDate = GetRandomDateInRange(monthStart, monthEnd)
        
        ' Get product with stock available on this date
        productSpecId = GetRandomProductWithStockOnDate(saleDate)
        
        If productSpecId > 0 Then
            ' CREATE THE DEBT
            debtId = CreateFakeDebt(debtorId, productSpecId, saleDate, totalAmount, expectedPaymentDate)
            
            If debtId > 0 Then
                totalDebts = totalDebts + 1
                Debug.Print "¦"
                Debug.Print "+--- Debt #" & debtId & " created"
                Debug.Print "¦    Sale Date: " & Format(saleDate, "dd/mm/yyyy")
                Debug.Print "¦    Amount: " & Format(totalAmount, "Currency")
                Debug.Print "¦    Expected: " & Format(expectedPaymentDate, "dd/mm/yyyy")
                
                ' DETERMINE PAYMENT SCENARIO
                ' 97% fully paid, 2% partially paid, 1% unpaid
                Dim randomValue As Double
                randomValue = Rnd
                
                If randomValue < 0.97 Then
                    ' Fully paid - scenarios 1-4
                    paymentScenario = Int((4 - 1 + 1) * Rnd + 1)
                ElseIf randomValue < 0.99 Then
                    ' Partially paid - scenario 6
                    paymentScenario = 6
                Else
                    ' Unpaid - scenario 5
                    paymentScenario = 5
                End If
                
                ' IMMEDIATELY GENERATE RETURNS FOR THIS DEBT
                returnsCreated = 0
                Call GenerateReturnsForDebt(debtId, totalAmount, saleDate, _
                                           expectedPaymentDate, finalEndDate, _
                                           paymentScenario, returnsCreated)
                totalReturns = totalReturns + returnsCreated
                
                ' VERIFY PAYMENT STATUS
                totalPaid = GetTotalPaidForDebt(debtId)
                
                If Abs(totalPaid - totalAmount) < 0.01 Then
                    ' Fully paid
                    paymentStatus = "? FULLY PAID"
                    totalFullyPaid = totalFullyPaid + 1
                ElseIf totalPaid > 0 Then
                    ' Partially paid
                    paymentStatus = "? PARTIAL (" & Format(totalPaid, "Currency") & " of " & Format(totalAmount, "Currency") & ")"
                    totalPartiallyPaid = totalPartiallyPaid + 1
                Else
                    ' Unpaid
                    paymentStatus = "? UNPAID"
                    totalUnpaid = totalUnpaid + 1
                End If
                
                Debug.Print "¦    Returns: " & returnsCreated
                Debug.Print "¦    Status: " & paymentStatus
                
            Else
                Debug.Print "¦    ? Failed to create debt"
            End If
        Else
            Debug.Print "¦    ? No product with stock found for " & Format(saleDate, "dd/mm/yyyy")
        End If
    Next j
    
    Debug.Print "+-------------------------------------"
    Debug.Print ""
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ProcessDebtorDebts: " & Err.description
End Sub

' Get random product that has stock available on a specific date
Private Function GetRandomProductWithStockOnDate(checkDate As Date) As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim productSpecId As Long
    Dim availableStock As Long
    Dim maxAttempts As Integer
    Dim attempts As Integer
    Dim totalProducts As Integer
    Dim randomPosition As Integer
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT product_spec_id FROM tbl_product_specs ORDER BY product_spec_id", dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        GetRandomProductWithStockOnDate = 0
        Exit Function
    End If
    
    rs.MoveLast
    totalProducts = rs.RecordCount
    rs.MoveFirst
    
    maxAttempts = totalProducts * 2
    attempts = 0
    productSpecId = 0
    
    Do While attempts < maxAttempts And productSpecId = 0
        ' Get random product
        randomPosition = Int(Rnd * totalProducts)
        
        If randomPosition > 0 Then
            rs.Move randomPosition
            If rs.EOF Then rs.MoveLast
        End If
        
        ' Check stock available up to this date
        availableStock = getCurrentStockQuantity(product_spec_id:=rs!product_spec_id, up_to_date:=checkDate)
        
        If availableStock > 0 Then
            productSpecId = rs!product_spec_id
        Else
            rs.MoveNext
            If rs.EOF Then rs.MoveFirst
        End If
        
        attempts = attempts + 1
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    GetRandomProductWithStockOnDate = productSpecId
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in GetRandomProductWithStockOnDate: " & Err.description
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    GetRandomProductWithStockOnDate = 0
End Function

' Create a fake debt with proper stock validation
Private Function CreateFakeDebt(debtorId As Long, productSpecId As Long, saleDate As Date, _
                                ByRef outAmount As Currency, ByRef outExpectedDate As Date) As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim quantity As Long
    Dim unitPrice As Currency
    Dim discount As Currency
    Dim daysUntilPayment As Integer
    Dim availableStock As Long
    Dim maxQty As Long
    Dim newDebtId As Long
    
    Set db = CurrentDb
    
    ' Check available stock on the sale date
    availableStock = getCurrentStockQuantity(product_spec_id:=productSpecId, up_to_date:=saleDate)
    
    If availableStock <= 0 Then
        CreateFakeDebt = 0
        Set db = Nothing
        Exit Function
    End If
    
    ' Determine max quantity we can sell
    maxQty = availableStock
    If maxQty > 10 Then maxQty = 10 ' Cap at 10 for realism
    
    ' Random quantity (1 to maxQty)
    If maxQty = 1 Then
        quantity = 1
    Else
        quantity = Int((maxQty - 1 + 1) * Rnd + 1)
    End If
    
    ' Get the suggested selling price
    unitPrice = getSuggestedSellingPrice(productSpecId)
    
    ' Discount is RARE (10% chance) and small (1-5%)
    If Rnd < 0.1 Then
        discount = CCur((quantity * unitPrice) * (Int((5 - 1 + 1) * Rnd + 1) / 100))
    Else
        discount = 0
    End If
    
    ' Random payment date (7 to 90 days after sale)
    daysUntilPayment = Int((90 - 7 + 1) * Rnd + 7)
    outExpectedDate = DateAdd("d", daysUntilPayment, saleDate)
    
    ' Calculate amount
    outAmount = CCur((quantity * unitPrice) - discount)
    
    ' Add record
    Set rs = db.OpenRecordset("tbl_debts", dbOpenDynaset)
    
    rs.AddNew
    rs!debtor_id = debtorId
    rs!product_spec_id = productSpecId
    rs!quantity = quantity
    rs!unit_price = unitPrice
    rs!discount = discount
    rs!sale_date = saleDate
    rs!expected_payment_date = outExpectedDate
    rs.upDate
    
    ' Get the new debt ID
    rs.Bookmark = rs.LastModified
    newDebtId = rs!sale_id
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    CreateFakeDebt = newDebtId
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in CreateFakeDebt: " & Err.description
    CreateFakeDebt = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Generate returns for a debt based on scenario
Private Sub GenerateReturnsForDebt(debtId As Long, totalAmount As Currency, _
                                   saleDate As Date, expectedDate As Date, EndDate As Date, _
                                   scenario As Integer, ByRef returnsCount As Long)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim paymentDate As Date
    Dim amountPaid As Currency
    Dim remainingAmount As Currency
    Dim numPayments As Integer
    Dim i As Integer
    Dim paymentMethodId As Long
    Dim baseDate As Date
    Dim totalDaysForPayments As Long
    Dim paymentPosition As Double
    Dim daysVariation As Integer
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_debt_returns", dbOpenDynaset)
    
    remainingAmount = totalAmount
    paymentMethodId = 1 ' Default payment method
    returnsCount = 0
    
    Select Case scenario
        Case 1 ' Paid in full - on time or early
            daysVariation = Int((10 - (-10) + 1) * Rnd + (-10)) ' -10 to +10 days
            paymentDate = DateAdd("d", daysVariation, expectedDate)
            If paymentDate < saleDate Then paymentDate = saleDate
            If paymentDate > EndDate Then paymentDate = EndDate
            
            rs.AddNew
            rs!debt_id = debtId
            rs!amount = totalAmount
            rs!return_date = paymentDate
            rs!payment_method_id = paymentMethodId
            If daysVariation < 0 Then
                rs!comment = "Paid in full (early)"
            ElseIf daysVariation <= 5 Then
                rs!comment = "Paid in full on time"
            Else
                rs!comment = "Paid in full (slightly late)"
            End If
            rs.upDate
            returnsCount = returnsCount + 1
            
        Case 2 ' Paid in full - significantly late (10-45 days overdue)
            paymentDate = DateAdd("d", Int((45 - 10 + 1) * Rnd + 10), expectedDate)
            If paymentDate > EndDate Then paymentDate = EndDate
            
            rs.AddNew
            rs!debt_id = debtId
            rs!amount = totalAmount
            rs!return_date = paymentDate
            rs!payment_method_id = paymentMethodId
            rs!comment = "Paid in full (overdue)"
            rs.upDate
            returnsCount = returnsCount + 1
            
        Case 3 ' Paid in portions - completed on time
            numPayments = Int((4 - 2 + 1) * Rnd + 2) ' 2-4 payments
            baseDate = DateAdd("d", Int((DateDiff("d", saleDate, expectedDate)) * 0.3), saleDate)
            totalDaysForPayments = DateDiff("d", baseDate, expectedDate)
            
            For i = 1 To numPayments
                ' Calculate payment amount - last payment gets exact remainder
                If i < numPayments Then
                    amountPaid = CCur(Int(totalAmount / numPayments))
                Else
                    amountPaid = remainingAmount
                End If
                
                ' Calculate payment date with spread
                paymentPosition = i / numPayments
                paymentDate = DateAdd("d", Int(totalDaysForPayments * paymentPosition), baseDate)
                paymentDate = DateAdd("d", Int((3 - (-3) + 1) * Rnd + (-3)), paymentDate)
                If paymentDate < saleDate Then paymentDate = saleDate
                If paymentDate > EndDate Then paymentDate = EndDate
                
                If amountPaid > 0 Then
                    rs.AddNew
                    rs!debt_id = debtId
                    rs!amount = amountPaid
                    rs!return_date = paymentDate
                    rs!payment_method_id = paymentMethodId
                    rs!comment = "Partial payment " & i & " of " & numPayments
                    rs.upDate
                    returnsCount = returnsCount + 1
                End If
                
                remainingAmount = remainingAmount - amountPaid
            Next i
            
        Case 4 ' Paid in portions - completed late
            numPayments = Int((4 - 2 + 1) * Rnd + 2) ' 2-4 payments
            Dim totalOverdueDays As Integer
            Dim daysBetweenPayments As Integer
            totalOverdueDays = Int((60 - 20 + 1) * Rnd + 20) ' 20-60 days total
            daysBetweenPayments = Int(totalOverdueDays / numPayments)
            If daysBetweenPayments < 3 Then daysBetweenPayments = 3
            
            For i = 1 To numPayments
                ' Calculate payment amount - last payment gets exact remainder
                If i < numPayments Then
                    amountPaid = CCur(Int(totalAmount / numPayments))
                Else
                    amountPaid = remainingAmount
                End If
                
                ' Each payment is progressively later
                paymentDate = DateAdd("d", daysBetweenPayments * i, expectedDate)
                paymentDate = DateAdd("d", Int((5 - (-2) + 1) * Rnd + (-2)), paymentDate)
                If paymentDate > EndDate Then paymentDate = EndDate
                
                If amountPaid > 0 Then
                    rs.AddNew
                    rs!debt_id = debtId
                    rs!amount = amountPaid
                    rs!return_date = paymentDate
                    rs!payment_method_id = paymentMethodId
                    rs!comment = "Partial payment " & i & " of " & numPayments & " (overdue)"
                    rs.upDate
                    returnsCount = returnsCount + 1
                End If
                
                remainingAmount = remainingAmount - amountPaid
            Next i
            
        Case 5 ' UNPAID - no payment at all (VERY RARE - 1%)
            ' No returns created
            returnsCount = 0
            
        Case 6 ' PARTIALLY PAID - some payment but not full (RARE - 2%)
            numPayments = Int((2 - 1 + 1) * Rnd + 1) ' 1-2 payments
            
            For i = 1 To numPayments
                ' Pay between 20-70% of remaining amount
                Dim paymentPercent As Double
                paymentPercent = 0.2 + (Rnd * 0.5) ' 20% to 70%
                amountPaid = CCur(Int(remainingAmount * paymentPercent))
                
                ' Random date after expected date
                paymentDate = DateAdd("d", Int(Rnd * 60), expectedDate)
                If paymentDate > EndDate Then paymentDate = EndDate
                
                If amountPaid > 0 And remainingAmount > 0 Then
                    If amountPaid > remainingAmount Then amountPaid = remainingAmount
                    
                    rs.AddNew
                    rs!debt_id = debtId
                    rs!amount = amountPaid
                    rs!return_date = paymentDate
                    rs!payment_method_id = paymentMethodId
                    rs!comment = "Partial payment (debt not fully settled)"
                    rs.upDate
                    returnsCount = returnsCount + 1
                    
                    remainingAmount = remainingAmount - amountPaid
                End If
            Next i
    End Select
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in GenerateReturnsForDebt: " & Err.description
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' Verify total amount paid for a debt
Private Function GetTotalPaidForDebt(debtId As Long) As Currency
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim total As Currency
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT SUM(amount) AS total_paid " & _
                              "FROM tbl_debt_returns WHERE debt_id = " & debtId, _
                              dbOpenSnapshot)
    
    total = 0
    If Not rs.EOF Then
        If Not IsNull(rs!total_paid) Then
            total = rs!total_paid
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    GetTotalPaidForDebt = total
    Exit Function
    
ErrorHandler:
    GetTotalPaidForDebt = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Get random date within a range
Private Function GetRandomDateInRange(startDate As Date, EndDate As Date) As Date
    Dim daysDiff As Long
    Dim randomDays As Long
    
    daysDiff = DateDiff("d", startDate, EndDate)
    If daysDiff > 0 Then
        randomDays = Int(Rnd * (daysDiff + 1))
        GetRandomDateInRange = DateAdd("d", randomDays, startDate)
    Else
        GetRandomDateInRange = startDate
    End If
End Function

