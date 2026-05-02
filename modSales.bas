Attribute VB_Name = "modSales"
Option Compare Database
Option Explicit
'modSales

Public Function DeleteAllTransactions(Optional ByVal ConfirmAsYes As Boolean = False) As Boolean
    ' Seriously: This function deletes ALL transactional data:
    '   - Sales, purchases, returns (inward/outward)
    '   - Drawings, office use withdrawals
    '   - Debts and debt returns (yes, the table is named tbl_debt_returns)
    '   - All payments, liability payment details
    '   - Prepayments, payment allocations
    '   - Payment obligations
    '   - Temporary tables
    '
    ' It keeps ALL master/setup data completely safe:
    '   Products, product specs, brands, categories, types, units,
    '   suppliers, debtors, assets, liabilities (items stay, only payment history cleared),
    '   expense types/items, payment methods, etc.
    '
    ' Param: ConfirmAsYes = True ? skips confirmation prompt
    '
    ' Returns: True if successful
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim response As VbMsgBoxResult
    
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    
    ' --------------------- SAFETY CONFIRMATION ---------------------
    If Not ConfirmAsYes Then
        response = MsgBox("SERIOUS WARNING:" & vbCrLf & vbCrLf & _
                          "This will PERMANENTLY delete every single transaction record in your database:" & vbCrLf & vbCrLf & _
                          "• All sales and purchases" & vbCrLf & _
                          "• All returns (inward and outward)" & vbCrLf & _
                          "• All drawings and office use" & vbCrLf & _
                          "• All debts and debt returns" & vbCrLf & _
                          "• All payments, prepayments, allocations and obligations" & vbCrLf & _
                          "• All liability payment details" & vbCrLf & vbCrLf & _
                          "Your entire setup (products, expense types, liability types, " & _
                          "payment methods, assets, etc.) will remain 100% untouched." & vbCrLf & vbCrLf & _
                          "This cannot be undone." & vbCrLf & vbCrLf & _
                          "Are you absolutely sure?", _
                          vbCritical + vbYesNo, "DELETE ALL TRANSACTIONS")
        
        If response = vbNo Then
            DeleteAllTransactions = False
            Exit Function
        End If
    End If
    
    ' --------------------- BEGIN TRANSACTION (correct DAO way) ---------------------
    ws.BeginTrans
    
    ' Delete in strict child-to-parent order to respect referential integrity
    
    db.Execute "DELETE FROM tbl_debt_returns", dbFailOnError                    ' debt repayments
    db.Execute "DELETE FROM tbl_return_inwards", dbFailOnError                  ' sales returns
    db.Execute "DELETE FROM tbl_return_outwards", dbFailOnError                 ' purchase returns
    db.Execute "DELETE FROM tbl_sales_office_use", dbFailOnError
    db.Execute "DELETE FROM tbl_drawings", dbFailOnError
    db.Execute "DELETE FROM tbl_debts", dbFailOnError                           ' credit sales
    db.Execute "DELETE FROM tbl_sales", dbFailOnError                           ' cash sales
    db.Execute "DELETE FROM tbl_purchase_details", dbFailOnError
    db.Execute "DELETE FROM tbl_purchases", dbFailOnError
    
    db.Execute "DELETE FROM tbl_payment_allocations", dbFailOnError
    db.Execute "DELETE FROM tbl_liability_payment_details", dbFailOnError
    db.Execute "DELETE FROM tbl_prepayments", dbFailOnError
    db.Execute "DELETE FROM tbl_payments", dbFailOnError
    db.Execute "DELETE FROM tbl_payment_obligations", dbFailOnError
    
    ' Temporary tables (ignore errors if they don't exist)
    On Error Resume Next
    db.Execute "DELETE FROM tbl_temp_COGS_Details", dbFailOnError
    db.Execute "DELETE FROM tmp_product_analysis", dbFailOnError
    On Error GoTo ErrorHandler
    
    db.Execute "UPDATE tbl_product_specs SET current_stock=0"
    ' --------------------- COMMIT ---------------------
    ws.CommitTrans
    
    MsgBox "All transaction data has been permanently and successfully deleted." & vbCrLf & vbCrLf & _
           "Your database now has a completely clean transaction history." & vbCrLf & _
           "All master data (products, configurations, etc.) is preserved.", _
           vbInformation, "Complete"
    
    DeleteAllTransactions = True
    Exit Function
    
ErrorHandler:
    ws.Rollback
    MsgBox "An error occurred (" & Err.Number & "): " & Err.description & vbCrLf & vbCrLf & _
           "No changes were made – everything was rolled back safely.", vbCritical, "Failed"
    DeleteAllTransactions = False
End Function

