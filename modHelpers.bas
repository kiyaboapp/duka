Attribute VB_Name = "modHelpers"
Option Compare Database
Option Explicit
'modHelpers

Public Function getDefaultUnit() As Long
    Dim answer As Long
    answer = Nz(DLookup("unit_id", "tbl_units", "unit_name LIKE 'piece*'"), 0)
    
    If answer = 0 Then
        answer = Nz(DLookup("unit_id", "tbl_units", "unit_abbr LIKE 'pc*'"), 0)
    End If
    
    If answer = 0 Then
        ' Create default unit record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_units", dbOpenDynaset)
        
        rs.AddNew
        rs!unit_name = "piece"
        rs!unit_abbr = "pc"
        answer = rs!unit_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultUnit = answer
End Function

Public Function getDefaultOfficeUse() As Long
    Dim answer As Long
    answer = Nz(DLookup("office_use_id", "tbl_office_use", "office_use LIKE '*Customer Care*'"), 0)

    
    If answer = 0 Then
        ' Create default unit record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_office_use", dbOpenDynaset)
        
        rs.AddNew
        rs!office_use = "Customer Care"
        answer = rs!office_use_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultOfficeUse = answer
End Function

Public Function getDefaultSpecValue() As Long
    Dim answer As Long
    answer = Nz(DLookup("spec_value_id", "tbl_spec_values", "spec_value = 'normal'"), 0)
    
    If answer = 0 Then
        answer = Nz(DLookup("spec_value_id", "tbl_spec_values", "spec_id = " & getDefaultSpecID()), 0)
    End If
    
    If answer = 0 Then
        ' Create default spec_value record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_spec_values", dbOpenDynaset)
        
        rs.AddNew
        rs!spec_id = getDefaultSpecID()
        rs!spec_value = "normal"
        answer = rs!spec_value_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultSpecValue = answer
End Function

Public Function getDefaultSpecID() As Long
    Dim answer As Long
    answer = Nz(DLookup("spec_id", "tbl_specs", "spec_name = 'normal'"), 0)
    
    If answer = 0 Then
        ' Create default spec record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_specs", dbOpenDynaset)
        
        rs.AddNew
        rs!spec_name = "normal"
        answer = rs!spec_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultSpecID = answer
End Function

Public Function getDefaultBrandID() As Long
    Dim answer As Long
    answer = Nz(DLookup("brand_id", "tbl_brands", "brand_name='No Brand'"), 0)
    
    If answer = 0 Then
        ' Create default brand record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_brands", dbOpenDynaset)
        
        rs.AddNew
        rs!brand_name = "No Brand"
        answer = rs!brand_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultBrandID = answer
End Function

Public Function getDefaultDrawingID() As Long
    Dim answer As Long
    answer = Nz(DLookup("drawing_category_id", "tbl_drawing_categories", "drawing_category='Business Owner'"), 0)
    
    If answer = 0 Then
        ' Create default brand record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_drawing_categories", dbOpenDynaset)
        
        rs.AddNew
        rs!drawing_category = "Business Owner"
        answer = rs!drawing_category_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultDrawingID = answer
End Function

Public Function getDefaultTypeID() As Long
    Dim answer As Long
    answer = Nz(DLookup("type_id", "tbl_types", "type_name='Others'"), 0)
    
    If answer = 0 Then
        ' Create default type record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_types", dbOpenDynaset)
        
        rs.AddNew
        rs!category_id = getDefaultCategoryID()
        rs!type_name = "Others"
        answer = rs!type_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultTypeID = answer
End Function


Public Function getDefaultCategoryID() As Long
    Dim answer As Long
    answer = Nz(DLookup("category_id", "tbl_categories", "category_name='Others Categories'"), 0)
    
    If answer = 0 Then
        ' Create default type record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_categories", dbOpenDynaset)
        
        rs.AddNew
        rs!category_name = "Others Categories"
        answer = rs!category_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultCategoryID = answer
End Function

Public Function getDefaultPaymentMethodID() As Long
    Dim answer As Long
    answer = Nz(DLookup("payment_method_id", "tbl_payment_methods", "payment_method='Cash'"), 0)
    
    If answer = 0 Then
        ' Create default type record
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbl_payment_methods", dbOpenDynaset)
        
        rs.AddNew
        rs!category_name = "Cash"
        answer = rs!payment_method_id
        rs.upDate
        
        
        rs.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    
    getDefaultPaymentMethodID = answer
End Function

Public Function getmsgTitle()
    getmsgTitle = "KIYABO DUKA"
End Function


Public Function getSuggestedSellingPrice(product_spec_id As Long) As Currency
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim suggestedPrice As Currency
    
    ' Default to 0 if nothing found
    getSuggestedSellingPrice = 0
    
    Set db = CurrentDb
    
    ' Query to get the suggested_selling_price from the most recent purchase record
    ' for the given product_spec_id
    sql = "SELECT TOP 1 suggested_selling_price " & _
          "FROM qry_purchases " & _
          "WHERE product_spec_id = " & product_spec_id & " " & _
          "ORDER BY purchase_date DESC;"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        If Not IsNull(rs!suggested_selling_price) Then
            suggestedPrice = rs!suggested_selling_price
            getSuggestedSellingPrice = suggestedPrice
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

