Attribute VB_Name = "modStockCRUD"
'================================================================================
' SALES & RETURNS CRUD SYSTEM WITH STOCK MANAGEMENT
' Core Parameters Order: product_spec_id, quantity, unit_price, then specifics
' NOTE: Uses DateValue() for accurate date comparisons (ignores time component)
'================================================================================

'================================================================================
' STOCK UPDATE FUNCTION
'================================================================================

Public Function UpdateProductsStock(Optional product_spec_id As Long = 0) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim oldStock As Variant
    Dim newStock As Variant
    
    Set db = CurrentDb
    
    If product_spec_id > 0 Then
        Set rs = db.OpenRecordset("SELECT current_stock FROM tbl_product_specs WHERE product_spec_id = " & product_spec_id, dbOpenSnapshot)
        If Not rs.EOF Then
            oldStock = Nz(rs!current_stock, 0)
        Else
            rs.Close
            UpdateProductsStock = False
            Exit Function
        End If
        rs.Close
        
        db.Execute "UPDATE tbl_product_specs ps " & _
                   "INNER JOIN qry_all_products q ON ps.product_spec_id = q.product_spec_id " & _
                   "SET ps.current_stock = q.stock " & _
                   "WHERE ps.product_spec_id = " & product_spec_id, dbFailOnError
        
        Set rs = db.OpenRecordset("SELECT current_stock FROM tbl_product_specs WHERE product_spec_id = " & product_spec_id, dbOpenSnapshot)
        If Not rs.EOF Then
            newStock = Nz(rs!current_stock, 0)
        End If
        rs.Close
        
        UpdateProductsStock = True
        
    Else
        db.Execute "UPDATE tbl_product_specs ps " & _
                   "INNER JOIN qry_all_products q ON ps.product_spec_id = q.product_spec_id " & _
                   "SET ps.current_stock = q.stock", dbFailOnError
        
        UpdateProductsStock = True
    End If
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    UpdateProductsStock = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' HELPER FUNCTION: Get Records by Date Range
'================================================================================

Public Function GetRecordsByDateRange( _
    tableName As String, _
    Optional product_spec_id As Long = 0, _
    Optional startDate As Variant = Null, _
    Optional EndDate As Variant = Null _
) As DAO.Recordset
    '------------------------------------------------------------------------
    ' Returns records filtered by product and date range
    ' Uses DateValue() to ignore time component in comparisons
    ' tableName: "tbl_sales", "tbl_drawings", etc.
    ' Returns: Recordset or Nothing
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim sql As String
    
    Set db = CurrentDb
    sql = "SELECT * FROM " & tableName & " WHERE 1=1"
    
    If product_spec_id > 0 Then
        sql = sql & " AND product_spec_id = " & product_spec_id
    End If
    
    If Not IsNull(startDate) Then
        sql = sql & " AND DateValue(sale_date) >= #" & Format(startDate, "mm/dd/yyyy") & "#"
    End If
    
    If Not IsNull(EndDate) Then
        sql = sql & " AND DateValue(sale_date) <= #" & Format(EndDate, "mm/dd/yyyy") & "#"
    End If
    
    Set GetRecordsByDateRange = db.OpenRecordset(sql, dbOpenSnapshot)
    Exit Function
    
ErrorHandler:
    Set GetRecordsByDateRange = Nothing
End Function

'================================================================================
' HELPER FUNCTION: Get Sum by Date Range
'================================================================================

Public Function GetSumByDateRange( _
    tableName As String, _
    sumField As String, _
    Optional product_spec_id As Long = 0, _
    Optional startDate As Variant = Null, _
    Optional EndDate As Variant = Null _
) As Currency
    '------------------------------------------------------------------------
    ' Returns sum of a field filtered by product and date range
    ' Uses DateValue() for accurate date comparisons
    ' Example: GetSumByDateRange("tbl_sales", "amount", 25, #12/1/2025#, #12/17/2025#)
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    sql = "SELECT Sum(" & sumField & ") AS total FROM " & tableName & " WHERE 1=1"
    
    If product_spec_id > 0 Then
        sql = sql & " AND product_spec_id = " & product_spec_id
    End If
    
    If Not IsNull(startDate) Then
        sql = sql & " AND DateValue(sale_date) >= #" & Format(startDate, "mm/dd/yyyy") & "#"
    End If
    
    If Not IsNull(EndDate) Then
        sql = sql & " AND DateValue(sale_date) <= #" & Format(EndDate, "mm/dd/yyyy") & "#"
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        GetSumByDateRange = Nz(rs!total, 0)
    Else
        GetSumByDateRange = 0
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetSumByDateRange = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' HELPER FUNCTION: Get Count by Date Range
'================================================================================

Public Function GetCountByDateRange( _
    tableName As String, _
    Optional product_spec_id As Long = 0, _
    Optional startDate As Variant = Null, _
    Optional EndDate As Variant = Null _
) As Long
    '------------------------------------------------------------------------
    ' Returns count of records filtered by product and date range
    ' Uses DateValue() for accurate date comparisons
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    sql = "SELECT Count(*) AS cnt FROM " & tableName & " WHERE 1=1"
    
    If product_spec_id > 0 Then
        sql = sql & " AND product_spec_id = " & product_spec_id
    End If
    
    If Not IsNull(startDate) Then
        sql = sql & " AND DateValue(sale_date) >= #" & Format(startDate, "mm/dd/yyyy") & "#"
    End If
    
    If Not IsNull(EndDate) Then
        sql = sql & " AND DateValue(sale_date) <= #" & Format(EndDate, "mm/dd/yyyy") & "#"
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        GetCountByDateRange = Nz(rs!cnt, 0)
    Else
        GetCountByDateRange = 0
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetCountByDateRange = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' TBL_SALES CRUD OPERATIONS
' Schema: sale_id, product_spec_id, quantity, unit_price, discount, amount,
'         sale_date, month_id, month_name, payment_method_id, notes
'================================================================================

Public Function CreateSale( _
    product_spec_id As Long, _
    quantity As Long, _
    unit_price As Currency, _
    Optional sale_date As Variant = Null, _
    Optional discount As Currency = 0, _
    Optional payment_method_id As Long = 0, _
    Optional notes As String = "" _
) As Long
    '------------------------------------------------------------------------
    ' Creates a new sale record
    ' Required: product_spec_id, quantity, unit_price
    ' Optional: sale_date (defaults to Now if not provided), discount,
    '           payment_method_id, notes
    ' Returns: sale_id of created record, 0 on failure
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim newSaleID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_sales", dbOpenDynaset)
    
    rs.AddNew
    rs!product_spec_id = product_spec_id
    rs!quantity = quantity
    rs!unit_price = unit_price
    
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If discount > 0 Then rs!discount = discount
    If payment_method_id > 0 Then rs!payment_method_id = payment_method_id
    If Len(notes) > 0 Then rs!notes = notes
    
    rs.upDate
    rs.Bookmark = rs.LastModified
    newSaleID = rs!sale_id
    rs.Close
    
    Call UpdateProductsStock(product_spec_id)
    
    CreateSale = newSaleID
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    CreateSale = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function ReadSale(sale_id As Long) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_sales WHERE sale_id = " & sale_id, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set ReadSale = rs
    Else
        rs.Close
        Set ReadSale = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set ReadSale = Nothing
End Function

Public Function UpdateSale( _
    sale_id As Long, _
    Optional product_spec_id As Variant = Null, _
    Optional quantity As Variant = Null, _
    Optional unit_price As Variant = Null, _
    Optional sale_date As Variant = Null, _
    Optional discount As Variant = Null, _
    Optional payment_method_id As Variant = Null, _
    Optional notes As Variant = Null _
) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim oldProductSpecID As Long
    Dim newProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_sales WHERE sale_id = " & sale_id, dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        UpdateSale = False
        Exit Function
    End If
    
    oldProductSpecID = rs!product_spec_id
    
    rs.Edit
    
    If Not IsNull(product_spec_id) Then rs!product_spec_id = product_spec_id
    If Not IsNull(quantity) Then rs!quantity = quantity
    If Not IsNull(unit_price) Then rs!unit_price = unit_price
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If Not IsNull(discount) Then rs!discount = discount
    If Not IsNull(payment_method_id) Then rs!payment_method_id = payment_method_id
    If Not IsNull(notes) Then rs!notes = notes
    
    rs.upDate
    newProductSpecID = rs!product_spec_id
    rs.Close
    
    If oldProductSpecID = newProductSpecID Then
        Call UpdateProductsStock(oldProductSpecID)
    Else
        Call UpdateProductsStock(oldProductSpecID)
        Call UpdateProductsStock(newProductSpecID)
    End If
    
    UpdateSale = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    UpdateSale = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function DeleteSale(sale_id As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim affectedProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT product_spec_id FROM tbl_sales WHERE sale_id = " & sale_id, dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close
        DeleteSale = False
        Exit Function
    End If
    
    affectedProductSpecID = rs!product_spec_id
    rs.Close
    
    db.Execute "DELETE FROM tbl_sales WHERE sale_id = " & sale_id, dbFailOnError
    
    Call UpdateProductsStock(affectedProductSpecID)
    
    DeleteSale = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    DeleteSale = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' TBL_RETURN_INWARDS CRUD OPERATIONS
' Schema: return_inward_id, sale_id, quantity, unit_price, amount, reason,
'         sale_date, month_id, month_name
'================================================================================

Public Function CreateReturnInward( _
    sale_id As Long, _
    quantity As Long, _
    unit_price As Currency, _
    Optional sale_date As Variant = Null, _
    Optional reason As String = "" _
) As Long
    '------------------------------------------------------------------------
    ' Creates a new return inward (sales return) record
    ' Required: sale_id, quantity, unit_price
    ' Optional: sale_date (defaults to Now if not provided), reason
    ' Returns: return_inward_id of created record, 0 on failure
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsProduct As DAO.Recordset
    Dim newReturnID As Long
    Dim productSpecId As Long
    
    Set db = CurrentDb
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_sales WHERE sale_id = " & sale_id, dbOpenSnapshot)
    If rsProduct.EOF Then
        rsProduct.Close
        CreateReturnInward = 0
        Exit Function
    End If
    productSpecId = rsProduct!product_spec_id
    rsProduct.Close
    
    Set rs = db.OpenRecordset("tbl_return_inwards", dbOpenDynaset)
    
    rs.AddNew
    rs!sale_id = sale_id
    rs!quantity = quantity
    rs!unit_price = unit_price
    
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If Len(reason) > 0 Then rs!reason = reason
    
    rs.upDate
    rs.Bookmark = rs.LastModified
    newReturnID = rs!return_inward_id
    rs.Close
    
    Call UpdateProductsStock(productSpecId)
    
    CreateReturnInward = newReturnID
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    CreateReturnInward = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function ReadReturnInward(return_inward_id As Long) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_return_inwards WHERE return_inward_id = " & return_inward_id, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set ReadReturnInward = rs
    Else
        rs.Close
        Set ReadReturnInward = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set ReadReturnInward = Nothing
End Function

Public Function UpdateReturnInward( _
    return_inward_id As Long, _
    Optional sale_id As Variant = Null, _
    Optional quantity As Variant = Null, _
    Optional unit_price As Variant = Null, _
    Optional sale_date As Variant = Null, _
    Optional reason As Variant = Null _
) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsProduct As DAO.Recordset
    Dim oldSaleID As Long
    Dim newSaleID As Long
    Dim oldProductSpecID As Long
    Dim newProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_return_inwards WHERE return_inward_id = " & return_inward_id, dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        UpdateReturnInward = False
        Exit Function
    End If
    
    oldSaleID = rs!sale_id
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_sales WHERE sale_id = " & oldSaleID, dbOpenSnapshot)
    oldProductSpecID = rsProduct!product_spec_id
    rsProduct.Close
    
    rs.Edit
    
    If Not IsNull(sale_id) Then rs!sale_id = sale_id
    If Not IsNull(quantity) Then rs!quantity = quantity
    If Not IsNull(unit_price) Then rs!unit_price = unit_price
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If Not IsNull(reason) Then rs!reason = reason
    
    newSaleID = rs!sale_id
    rs.upDate
    rs.Close
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_sales WHERE sale_id = " & newSaleID, dbOpenSnapshot)
    newProductSpecID = rsProduct!product_spec_id
    rsProduct.Close
    
    If oldProductSpecID = newProductSpecID Then
        Call UpdateProductsStock(oldProductSpecID)
    Else
        Call UpdateProductsStock(oldProductSpecID)
        Call UpdateProductsStock(newProductSpecID)
    End If
    
    UpdateReturnInward = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    UpdateReturnInward = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function DeleteReturnInward(return_inward_id As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsProduct As DAO.Recordset
    Dim saleID As Long
    Dim affectedProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT sale_id FROM tbl_return_inwards WHERE return_inward_id = " & return_inward_id, dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close
        DeleteReturnInward = False
        Exit Function
    End If
    
    saleID = rs!sale_id
    rs.Close
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_sales WHERE sale_id = " & saleID, dbOpenSnapshot)
    affectedProductSpecID = rsProduct!product_spec_id
    rsProduct.Close
    
    db.Execute "DELETE FROM tbl_return_inwards WHERE return_inward_id = " & return_inward_id, dbFailOnError
    
    Call UpdateProductsStock(affectedProductSpecID)
    
    DeleteReturnInward = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    DeleteReturnInward = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' TBL_RETURN_OUTWARDS CRUD OPERATIONS
' Schema: return_outward_id, purchase_detail_id, quantity, unit_price, amount,
'         reason, sale_date, month_id, month_name
'================================================================================

Public Function CreateReturnOutward( _
    purchase_detail_id As Long, _
    quantity As Long, _
    unit_price As Currency, _
    Optional sale_date As Variant = Null, _
    Optional reason As String = "" _
) As Long
    '------------------------------------------------------------------------
    ' Creates a new return outward (purchase return) record
    ' Required: purchase_detail_id, quantity, unit_price
    ' Optional: sale_date (defaults to Now if not provided), reason
    ' Returns: return_outward_id of created record, 0 on failure
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsProduct As DAO.Recordset
    Dim newReturnID As Long
    Dim productSpecId As Long
    
    Set db = CurrentDb
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_purchase_details WHERE purchase_detail_id = " & purchase_detail_id, dbOpenSnapshot)
    If rsProduct.EOF Then
        rsProduct.Close
        CreateReturnOutward = 0
        Exit Function
    End If
    productSpecId = rsProduct!product_spec_id
    rsProduct.Close
    
    Set rs = db.OpenRecordset("tbl_return_outwards", dbOpenDynaset)
    
    rs.AddNew
    rs!purchase_detail_id = purchase_detail_id
    rs!quantity = quantity
    rs!unit_price = unit_price
    
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If Len(reason) > 0 Then rs!reason = reason
    
    rs.upDate
    rs.Bookmark = rs.LastModified
    newReturnID = rs!return_outward_id
    rs.Close
    
    Call UpdateProductsStock(productSpecId)
    
    CreateReturnOutward = newReturnID
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    CreateReturnOutward = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function ReadReturnOutward(return_outward_id As Long) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_return_outwards WHERE return_outward_id = " & return_outward_id, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set ReadReturnOutward = rs
    Else
        rs.Close
        Set ReadReturnOutward = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set ReadReturnOutward = Nothing
End Function

Public Function UpdateReturnOutward( _
    return_outward_id As Long, _
    Optional purchase_detail_id As Variant = Null, _
    Optional quantity As Variant = Null, _
    Optional unit_price As Variant = Null, _
    Optional sale_date As Variant = Null, _
    Optional reason As Variant = Null _
) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsProduct As DAO.Recordset
    Dim oldPurchaseDetailID As Long
    Dim newPurchaseDetailID As Long
    Dim oldProductSpecID As Long
    Dim newProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_return_outwards WHERE return_outward_id = " & return_outward_id, dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        UpdateReturnOutward = False
        Exit Function
    End If
    
    oldPurchaseDetailID = rs!purchase_detail_id
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_purchase_details WHERE purchase_detail_id = " & oldPurchaseDetailID, dbOpenSnapshot)
    oldProductSpecID = rsProduct!product_spec_id
    rsProduct.Close
    
    rs.Edit
    
    If Not IsNull(purchase_detail_id) Then rs!purchase_detail_id = purchase_detail_id
    If Not IsNull(quantity) Then rs!quantity = quantity
    If Not IsNull(unit_price) Then rs!unit_price = unit_price
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If Not IsNull(reason) Then rs!reason = reason
    
    newPurchaseDetailID = rs!purchase_detail_id
    rs.upDate
    rs.Close
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_purchase_details WHERE purchase_detail_id = " & newPurchaseDetailID, dbOpenSnapshot)
    newProductSpecID = rsProduct!product_spec_id
    rsProduct.Close
    
    If oldProductSpecID = newProductSpecID Then
        Call UpdateProductsStock(oldProductSpecID)
    Else
        Call UpdateProductsStock(oldProductSpecID)
        Call UpdateProductsStock(newProductSpecID)
    End If
    
    UpdateReturnOutward = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    UpdateReturnOutward = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function DeleteReturnOutward(return_outward_id As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsProduct As DAO.Recordset
    Dim purchaseDetailID As Long
    Dim affectedProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT purchase_detail_id FROM tbl_return_outwards WHERE return_outward_id = " & return_outward_id, dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close
        DeleteReturnOutward = False
        Exit Function
    End If
    
    purchaseDetailID = rs!purchase_detail_id
    rs.Close
    
    Set rsProduct = db.OpenRecordset("SELECT product_spec_id FROM tbl_purchase_details WHERE purchase_detail_id = " & purchaseDetailID, dbOpenSnapshot)
    affectedProductSpecID = rsProduct!product_spec_id
    rsProduct.Close
    
    db.Execute "DELETE FROM tbl_return_outwards WHERE return_outward_id = " & return_outward_id, dbFailOnError
    
    Call UpdateProductsStock(affectedProductSpecID)
    
    DeleteReturnOutward = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    DeleteReturnOutward = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' TBL_DRAWINGS CRUD OPERATIONS
' Schema: drawing_id, drawing_category_id, product_spec_id, quantity, unit_price,
'         discount, amount, sale_date, month_id, month_name
'================================================================================

Public Function CreateDrawing( _
    product_spec_id As Long, _
    quantity As Long, _
    unit_price As Currency, _
    drawing_category_id As Long, _
    Optional sale_date As Variant = Null, _
    Optional discount As Currency = 0, _
    Optional notes As String = "" _
) As Long
    '------------------------------------------------------------------------
    ' Creates a new drawing record
    ' Required: product_spec_id, quantity, unit_price, drawing_category_id
    ' Optional: sale_date (defaults to Now if not provided), discount
    ' Returns: drawing_id of created record, 0 on failure
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim newDrawingID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_drawings", dbOpenDynaset)
    
    rs.AddNew
    rs!product_spec_id = product_spec_id
    rs!quantity = quantity
    rs!unit_price = unit_price
    rs!drawing_category_id = drawing_category_id
    rs!notes = notes
    
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If discount > 0 Then rs!discount = discount
    
    rs.upDate
    rs.Bookmark = rs.LastModified
    newDrawingID = rs!drawing_id
    rs.Close
    
    Call UpdateProductsStock(product_spec_id)
    
    CreateDrawing = newDrawingID
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    CreateDrawing = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function ReadDrawing(drawing_id As Long) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_drawings WHERE drawing_id = " & drawing_id, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set ReadDrawing = rs
    Else
        rs.Close
        Set ReadDrawing = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set ReadDrawing = Nothing
End Function

Public Function UpdateDrawing( _
    drawing_id As Long, _
    Optional product_spec_id As Variant = Null, _
    Optional quantity As Variant = Null, _
    Optional unit_price As Variant = Null, _
    Optional drawing_category_id As Variant = Null, _
    Optional sale_date As Variant = Null, _
    Optional discount As Variant = Null, _
    Optional notes As Variant = Null _
) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim oldProductSpecID As Long
    Dim newProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_drawings WHERE drawing_id = " & drawing_id, dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        UpdateDrawing = False
        Exit Function
    End If
    
    If IsNull(notes) Then notes = ""
    oldProductSpecID = rs!product_spec_id
    
    rs.Edit
    
    If Not IsNull(product_spec_id) Then rs!product_spec_id = product_spec_id
    If Not IsNull(quantity) Then rs!quantity = quantity
    If Not IsNull(unit_price) Then rs!unit_price = unit_price
    If Not IsNull(drawing_category_id) Then rs!drawing_category_id = drawing_category_id
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If Not IsNull(discount) Then rs!discount = discount
    If Not IsNull(notes) Then rs!notes = notes
    rs.upDate
    newProductSpecID = rs!product_spec_id
    rs.Close
    
    If oldProductSpecID = newProductSpecID Then
        Call UpdateProductsStock(oldProductSpecID)
    Else
        Call UpdateProductsStock(oldProductSpecID)
        Call UpdateProductsStock(newProductSpecID)
    End If
    
    UpdateDrawing = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    UpdateDrawing = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function DeleteDrawing(drawing_id As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim affectedProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT product_spec_id FROM tbl_drawings WHERE drawing_id = " & drawing_id, dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close
        DeleteDrawing = False
        Exit Function
    End If
    
    affectedProductSpecID = rs!product_spec_id
    rs.Close
    
    db.Execute "DELETE FROM tbl_drawings WHERE drawing_id = " & drawing_id, dbFailOnError
    
    Call UpdateProductsStock(affectedProductSpecID)
    
    DeleteDrawing = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    DeleteDrawing = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' TBL_SALES_OFFICE_USE CRUD OPERATIONS
' Schema: sale_id, product_spec_id, quantity, unit_price, discount, amount,
'         sale_date, month_id, month_name, reason, office_use_id
'================================================================================

Public Function CreateSalesOfficeUse( _
    product_spec_id As Long, _
    quantity As Long, _
    unit_price As Currency, _
    Optional sales_sale_id As Long = 0, _
    Optional sale_date As Variant = Null, _
    Optional office_use_id As Long = 0, _
    Optional discount As Currency = 0, _
    Optional reason As String = "" _
) As Long
    '------------------------------------------------------------------------
    ' Creates a new sales office use record
    ' Required: product_spec_id, quantity, unit_price
    ' Optional: sale_date (defaults to Now if not provided), office_use_id,
    '           discount, reason
    ' Returns: sale_id of created record, 0 on failure
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim newSaleID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_sales_office_use", dbOpenDynaset)
    
    rs.AddNew
    rs!product_spec_id = product_spec_id
    rs!quantity = quantity
    rs!unit_price = unit_price
    
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If office_use_id > 0 Then rs!office_use_id = office_use_id
    If discount > 0 Then rs!discount = discount
    If Len(reason) > 0 Then rs!reason = reason
    If sales_sale_id > 0 Then rs!sales_sale_id = sales_sale_id
    
    rs.upDate
    rs.Bookmark = rs.LastModified
    newSaleID = rs!sale_id
    rs.Close
    
    Call UpdateProductsStock(product_spec_id)
    
    CreateSalesOfficeUse = newSaleID
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    CreateSalesOfficeUse = 0
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function ReadSalesOfficeUse(sale_id As Long) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_sales_office_use WHERE sale_id = " & sale_id, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set ReadSalesOfficeUse = rs
    Else
        rs.Close
        Set ReadSalesOfficeUse = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set ReadSalesOfficeUse = Nothing
End Function

Public Function UpdateSalesOfficeUse( _
    sale_id As Long, _
    Optional sales_sale_id As Variant = Null, _
    Optional product_spec_id As Variant = Null, _
    Optional quantity As Variant = Null, _
    Optional unit_price As Variant = Null, _
    Optional sale_date As Variant = Null, _
    Optional office_use_id As Variant = Null, _
    Optional discount As Variant = Null, _
    Optional reason As Variant = Null _
) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim oldProductSpecID As Long
    Dim newProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbl_sales_office_use WHERE sale_id = " & sale_id, dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        UpdateSalesOfficeUse = False
        Exit Function
    End If
    
    oldProductSpecID = rs!product_spec_id
    
    rs.Edit
    
    If Not IsNull(product_spec_id) Then rs!product_spec_id = product_spec_id
    If Not IsNull(quantity) Then rs!quantity = quantity
    If Not IsNull(unit_price) Then rs!unit_price = unit_price
    If Not IsNull(sale_date) Then rs!sale_date = sale_date
    If Not IsNull(office_use_id) Then rs!office_use_id = office_use_id
    If Not IsNull(discount) Then rs!discount = discount
    If Not IsNull(reason) Then rs!reason = reason
    If Not IsNull(sales_sale_id) Then rs!sales_sale_id = sales_sale_id
    
    rs.upDate
    newProductSpecID = rs!product_spec_id
    rs.Close
    
    If oldProductSpecID = newProductSpecID Then
        Call UpdateProductsStock(oldProductSpecID)
    Else
        Call UpdateProductsStock(oldProductSpecID)
        Call UpdateProductsStock(newProductSpecID)
    End If
    
    UpdateSalesOfficeUse = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    UpdateSalesOfficeUse = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Public Function DeleteSalesOfficeUse(sale_id As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim affectedProductSpecID As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT product_spec_id FROM tbl_sales_office_use WHERE sale_id = " & sale_id, dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close
        DeleteSalesOfficeUse = False
        Exit Function
    End If
    
    affectedProductSpecID = rs!product_spec_id
    rs.Close
    
    db.Execute "DELETE FROM tbl_sales_office_use WHERE sale_id = " & sale_id, dbFailOnError
    
    Call UpdateProductsStock(affectedProductSpecID)
    
    DeleteSalesOfficeUse = True
    
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    DeleteSalesOfficeUse = False
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

'================================================================================
' USAGE EXAMPLES
'================================================================================

'Example 1: Create sale with default date (Now)
'saleID = CreateSale(product_spec_id:=1, quantity:=10, unit_price:=50.00)

'Example 2: Create sale with custom date
'saleID = CreateSale(product_spec_id:=1, quantity:=10, unit_price:=50.00, _
'                    sale_date:=#12/25/2024#, discount:=5.00)

'Example 3: Create return with custom date
'returnID = CreateReturnInward(sale_id:=5, quantity:=2, unit_price:=50.00, _
'                              sale_date:=#12/26/2024#, reason:="Defective")

'Example 4: Create drawing with default date
'drawingID = CreateDrawing(product_spec_id:=1, quantity:=5, unit_price:=50.00, _
'                          drawing_category_id:=1)

'Example 5: Update with custom date
'success = UpdateSale(sale_id:=5, sale_date:=#1/1/2025#, quantity:=15)

'================================================================================
' END OF MODULE
'================================================================================

