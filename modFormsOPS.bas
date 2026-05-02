Attribute VB_Name = "modFormsOPS"
Option Compare Database
Option Explicit
'modFormsOPS

Public Sub OpenSalesForm( _
    Optional sale_id As Long = 0, _
    Optional product_spec_id As Long = 0, _
    Optional quantity As Long = 0, _
    Optional unit_price As Currency = 0, _
    Optional discount As Currency = 0, _
    Optional payment_method_id As Long = 0, _
    Optional notes As String = "")

    On Error GoTo ErrorHandler
    
    DoCmd.OpenForm "frm_sales_single", acNormal
    
    With Forms!frm_sales_single
    
        ' Added this line to set the sale_id control
        If sale_id > 0 Then
            !sale_id = sale_id          ' Use your exact control name here
        End If
        
        If product_spec_id > 0 Then
            !product_spec_id = product_spec_id
        End If
        
        If quantity > 0 Then
            !quantity = quantity
        End If
        
        If unit_price > 0 Then
            !unit_price = unit_price
        End If
        
        If discount > 0 Then
            !discount = discount
        End If
        
        If payment_method_id > 0 Then
            !payment_method_id = payment_method_id
        End If
        
        If notes <> "" Then
            !notes = notes
        End If
        
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.description, vbCritical

End Sub


Public Sub OpenOfficUseForm( _
    Optional sale_id As Long = 0, _
    Optional office_use_id As Long = 0, _
    Optional sales_sale_id As Long = 0, _
    Optional product_spec_id As Long = 0, _
    Optional quantity As Long = 0, _
    Optional discount As Currency = 0, _
    Optional date_sold As Variant = Null, _
    Optional payment_method_id As Long = 0, _
    Optional reason As String = "", _
    Optional unit_price As Currency = 0)

    On Error GoTo ErrorHandler
    
    DoCmd.OpenForm "frm_sales_office_use_single", acNormal
    
    With Forms!frm_sales_office_use_single
        
        If sale_id > 0 Then
            !sale_id = sale_id
        End If
        
        If office_use_id > 0 Then
            !office_use_id = office_use_id
        End If
        
        If sales_sale_id > 0 Then
            !sales_sale_id = sales_sale_id
        End If
        
        If product_spec_id > 0 Then
            !product_spec_id = product_spec_id
        End If
        
        If quantity > 0 Then
            !quantity = quantity
        End If
        
        If discount > 0 Then
            !discount = discount
        End If
        
        If Not IsNull(date_sold) Then
            !date_sold = date_sold
        End If
        
        If payment_method_id > 0 Then
            !payment_method_id = payment_method_id
        End If
        
        If reason <> "" Then
            !reason = reason
        End If
        
        If unit_price = 0 Then unit_price = getSuggestedSellingPrice(product_spec_id)
        !unit_price = unit_price
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.description, vbCritical
End Sub

Public Sub OpenDrawingForm( _
    Optional drawing_id As Long = 0, _
    Optional product_spec_id As Long = 0, _
    Optional quantity As Long = 0, _
    Optional unit_price As Currency = 0, _
    Optional discount As Currency = 0, _
    Optional sale_date As Variant = Null, _
    Optional notes As String = "")

    On Error GoTo ErrorHandler
    
    DoCmd.OpenForm "frm_drawings_single", acNormal
    
    With Forms!frm_drawings_single
        
        If drawing_id > 0 Then
            !drawing_id = drawing_id
        End If
        
        If product_spec_id > 0 Then
            !product_spec_id = product_spec_id
        End If
        
        If quantity > 0 Then
            !quantity = quantity
        End If
        
        If unit_price > 0 Then
            !unit_price = unit_price
        Else
            !unit_price = getSuggestedSellingPrice(product_spec_id)
        End If
        
        If discount > 0 Then
            !discount = discount
        End If
        
        If Not IsNull(sale_date) Then
            !sale_date = sale_date
        End If
        
        If notes <> "" Then
            !notes = notes
        End If
        
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.description, vbCritical
End Sub
                    

