# Kiyabo Duka — Complete Development Master Document
**Database:** Kiyabo Duka v0.031.accdb  
**Platform:** Microsoft Access (Windows)  
**Prepared:** May 2026  
**Status:** Active Development — v0.031  

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Current State Audit](#2-current-state-audit)
3. [Architecture & Design Philosophy](#3-architecture--design-philosophy)
4. [frm_master — The Dashboard (Build This First)](#4-frm_master--the-dashboard-build-this-first)
5. [Forms — Complete Specification](#5-forms--complete-specification)
6. [Reports — Complete Specification](#6-reports--complete-specification)
7. [VBA Module Reference](#7-vba-module-reference)
8. [Schema Improvements](#8-schema-improvements)
9. [Build Roadmap & Priority Order](#9-build-roadmap--priority-order)
10. [Speed-Up Tools & Boilerplates](#10-speed-up-tools--boilerplates)
11. [Access VBA Boilerplate Library](#11-access-vba-boilerplate-library)
12. [Online Tools & Resources](#12-online-tools--resources)

---

## 1. Project Overview

Kiyabo Duka is a **retail shop management system** built in Microsoft Access. It tracks:

- Product catalog (categories → types → products → specs/variants)
- Purchasing from suppliers
- Sales (cash, credit/debt, office use, drawings/withdrawals)
- Debt collection (credit sales and repayments)
- Expenses (recurring and one-off)
- Liabilities (loans, payables)
- Assets (fixed assets register)
- Financial reporting (P&L, cash flow, balance sheet)

The data model is **mature and well-designed**. The gap is almost entirely in the UI layer — forms and reports. Think of it as: the engine is built, but the dashboard and instruments are missing.

---

## 2. Current State Audit

### 2.1 Tables (33 total)

| Table | Records | Status |
|---|---|---|
| tbl_sales | 4,020 | ✅ Active |
| tbl_purchase_details | 357 | ✅ Active |
| tbl_debt_returns | 191 | ✅ Active — no form |
| tbl_payment_obligations | 102 | ✅ Active — no form |
| tbl_debts | 100 | ✅ Active — no form |
| tbl_product_specs | 90 | ✅ Active |
| tbl_products | 123 | ✅ Active |
| tbl_purchases | 74 | ✅ Active — no header form |
| tbl_payments | 23 | ⚠️ Low — no form |
| tbl_assets | 8 | ⚠️ Low — no form |
| tbl_drawings | 0 | ❌ Unused — no form |
| tbl_return_inwards | 0 | ❌ Unused — no form |
| tbl_liability_payment_details | 0 | ❌ Unused — no form |

### 2.2 Forms (10 existing)

| Form | Purpose | Gap |
|---|---|---|
| frm_sales | Sales entry | Missing payment session grouping |
| frm_sales_single | Single sale view | OK |
| frm_sales_office_use_single | Office use entry | OK |
| frm_sales_office_use_datasheet | Office use list | OK |
| frm_products_list | Product listing | Missing barcode field |
| frm_product_specs | Spec entry | OK |
| frm_purchase_details | Purchase line entry | No parent purchase header |
| frm_drawings_single | Drawing entry | No list/main form |
| frm_report | Report launcher | Incomplete |
| frm_search_advanced | Search | Scope unknown |

### 2.3 Reports (2 existing — both are drafts)

| Report | Status |
|---|---|
| rpt_sample | Placeholder only |
| rpt_try | Experimental only |

**Zero production reports exist.**

### 2.4 VBA Modules (9)

| Module | Likely Purpose |
|---|---|
| modAccountings | P&L, financial calculations |
| modExpenses | Expense obligation generation |
| modFake | Test/seed data |
| modFormsOPS | Form helper functions |
| modHelpers | Utility functions |
| modPayables | Payables processing |
| modRATE | Expense rate calculations |
| modSales | Sales processing logic |
| modSchema | Schema/DDL helpers |
| modStats | Statistical summaries |
| modStockCRUD | Stock level updates |

---

## 3. Architecture & Design Philosophy

### 3.1 Navigation Model

```
frm_master (Dashboard — Landing Page)
├── SALES
│   ├── frm_sales_pos         (Point of Sale entry)
│   ├── frm_sales_list        (Sales history + search)
│   └── frm_debts             (Credit sales + collections)
├── STOCK
│   ├── frm_purchases         (Purchase orders)
│   ├── frm_products_list     (Product catalog)
│   └── frm_returns           (Return inwards / outwards)
├── EXPENSES
│   ├── frm_expenses_list     (Expense items + obligations)
│   └── frm_payments_list     (Payment recording)
├── FINANCE
│   ├── frm_assets            (Fixed assets register)
│   ├── frm_liabilities       (Loans & liabilities)
│   └── frm_drawings          (Owner withdrawals)
└── REPORTS
    └── frm_reports           (Report launcher hub)
```

### 3.2 Form Design Standards

Every form in the system should follow these rules:

**Header Section (every form)**
- Shop name / logo top-left
- Form title centered (large, bold)
- Date/time display top-right
- Navigation buttons: [🏠 Home] [← Back] [Close]

**Body Section**
- Label width: consistent 120px minimum
- Input fields: use combo boxes for all FK lookups
- Required fields: red asterisk (*) beside label
- Currency fields: formatted as `#,##0.00` with "TZS" label

**Footer Section**
- [➕ New Record] [💾 Save] [🗑️ Delete] [🖨️ Print] [❌ Close]
- Record counter: "Record X of Y"
- Status bar message area

**Color Scheme (suggested)**
```
Header background:  #1E3A5F  (dark navy)
Header text:        #FFFFFF
Form background:    #F5F5F5
Section headers:    #2E6DA4  (medium blue)
Input backgrounds:  #FFFFFF
Required fields:    #FFF3CD  (light yellow)
Buttons:            #2E6DA4 with white text
Save button:        #28A745 (green)
Delete button:      #DC3545 (red)
```

### 3.3 Naming Conventions

| Object | Convention | Example |
|---|---|---|
| Forms | frm_[module]_[action] | frm_sales_pos |
| Reports | rpt_[subject]_[type] | rpt_sales_daily |
| Queries | qry_[purpose] | qry_sales_by_month |
| Macros | mcr_[action] | mcr_open_sales |
| Subforms | sub_[parent]_[child] | sub_purchases_details |

---

## 4. frm_master — The Dashboard (Build This First)

The master dashboard is the **first thing users see** when they open the database. It replaces the default Access navigation pane with a professional, touch-friendly interface.

### 4.1 Layout Design

```
┌─────────────────────────────────────────────────────────────────┐
│  🏪  KIYABO DUKA                          📅 Saturday, 3 May 2026│
│      Duka Management System                🕐 14:32              │
├─────────────────────────────────────────────────────────────────┤
│                                                                   │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐  │
│  │                 │  │                 │  │                 │  │
│  │   💰  SALES     │  │  📦  STOCK      │  │  💳  EXPENSES   │  │
│  │                 │  │                 │  │                 │  │
│  │  [New Sale]     │  │  [Purchases]    │  │  [View Bills]   │  │
│  │  [Sales History]│  │  [Products]     │  │  [Pay Expense]  │  │
│  │  [Debtors]      │  │  [Returns]      │  │                 │  │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘  │
│                                                                   │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐  │
│  │                 │  │                 │  │                 │  │
│  │  📊  FINANCE    │  │  🖨️  REPORTS    │  │  ⚙️  SETTINGS  │  │
│  │                 │  │                 │  │                 │  │
│  │  [Assets]       │  │  [Daily Sales]  │  │  [Products]     │  │
│  │  [Liabilities]  │  │  [Stock Levels] │  │  [Suppliers]    │  │
│  │  [Drawings]     │  │  [Debtor Report]│  │  [Debtors]      │  │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘  │
│                                                                   │
├─────────────────────────────────────────────────────────────────┤
│  📈 TODAY: Sales TZS 0  │  ⚠️ Low Stock: 0 items  │  💰 Debts: 0│
└─────────────────────────────────────────────────────────────────┘
```

### 4.2 Dashboard Metrics (Bottom Status Bar)

These 5 KPIs update automatically on form load:

| Metric | Query Source | VBA Function |
|---|---|---|
| Today's Sales | `SELECT Sum(amount) FROM tbl_sales WHERE sale_date >= Date()` | `GetTodaySales()` |
| This Month Sales | `SELECT Sum(amount) FROM tbl_sales WHERE month_id = Month(Date())` | `GetMonthSales()` |
| Low Stock Items | `SELECT Count(*) FROM tbl_product_specs WHERE current_stock <= reorder_level` | `GetLowStockCount()` |
| Total Debt Outstanding | `SELECT Sum(amount - [amount paid back]) FROM tbl_debts` | `GetTotalDebt()` |
| Unpaid Obligations | `SELECT Count(*) FROM tbl_payment_obligations WHERE payment_status <> 'Paid'` | `GetUnpaidObligations()` |

### 4.3 VBA Code for frm_master

```vba
' ============================================================
' frm_master — On Load Event
' ============================================================
Private Sub Form_Load()
    ' Set title bar
    Me.Caption = "Kiyabo Duka — Dashboard"
    
    ' Display date and time
    Me.lblDate.Caption = Format(Date, "dddd, d mmmm yyyy")
    Me.lblTime.Caption = Format(Time, "hh:mm")
    
    ' Load KPI metrics
    Call RefreshDashboardStats
    
    ' Start timer for clock (1000ms = 1 second)
    Me.TimerInterval = 1000
End Sub

Private Sub Form_Timer()
    Me.lblTime.Caption = Format(Now(), "hh:mm:ss")
End Sub

Private Sub RefreshDashboardStats()
    On Error Resume Next
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' Today's sales
    Me.lblTodaySales.Caption = "TZS " & Format(Nz(DSum("amount", "tbl_sales", _
        "sale_date >= #" & Format(Date, "mm/dd/yyyy") & "#"), 0), "#,##0.00")
    
    ' Low stock count
    Me.lblLowStock.Caption = Nz(DCount("*", "tbl_product_specs", _
        "current_stock <= reorder_level"), 0) & " items"
    
    ' Outstanding debt (approximate — adjust query to your debt calc)
    ' Total debts minus total returns
    Dim dblDebts As Double
    Dim dblReturns As Double
    dblDebts = Nz(DSum("amount", "tbl_debts"), 0)
    dblReturns = Nz(DSum("amount", "tbl_debt_returns"), 0)
    Me.lblOutstandingDebt.Caption = "TZS " & Format(dblDebts - dblReturns, "#,##0.00")
    
    Set db = Nothing
End Sub

' ============================================================
' Navigation Buttons
' ============================================================
Private Sub btnNewSale_Click()
    DoCmd.OpenForm "frm_sales_pos", acNormal
End Sub

Private Sub btnSalesHistory_Click()
    DoCmd.OpenForm "frm_sales_list", acNormal
End Sub

Private Sub btnDebtors_Click()
    DoCmd.OpenForm "frm_debtors", acNormal
End Sub

Private Sub btnPurchases_Click()
    DoCmd.OpenForm "frm_purchases", acNormal
End Sub

Private Sub btnProducts_Click()
    DoCmd.OpenForm "frm_products_list", acNormal
End Sub

Private Sub btnReports_Click()
    DoCmd.OpenForm "frm_reports_hub", acNormal
End Sub

Private Sub btnClose_Click()
    If MsgBox("Are you sure you want to close Kiyabo Duka?", _
              vbYesNo + vbQuestion, "Close Application") = vbYes Then
        Application.Quit acSaveYes
    End If
End Sub
```

---

## 5. Forms — Complete Specification

---

### FORM 1: frm_sales_pos (Point of Sale)

**Purpose:** Primary daily sales entry — one item per transaction (current model).  
**Tables:** tbl_sales, tbl_product_specs, tbl_products, tbl_payment_methods  
**Priority:** P0 — already partially exists as frm_sales

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  💰 NEW SALE                    📅 03/05/2026  🕐 14:32 │
├─────────────────────────────────────────────────────────┤
│  Product:    [▼ Select Product + Spec         ] [🔍]    │
│  Quantity:   [    1    ]  Unit: [ pcs ]                 │
│  Unit Price: [ 0.00    ]  (auto-filled from spec)       │
│  Discount:   [ 0.00    ]                                │
│  ─────────────────────────────────────────────────────  │
│  AMOUNT:     TZS  0.00          (calculated)            │
│  ─────────────────────────────────────────────────────  │
│  Payment:    [▼ Cash / Mobile / Card / Credit ]         │
│  Notes:      [                               ]          │
├─────────────────────────────────────────────────────────┤
│  [➕ Save & New]   [💾 Save]   [🗑️ Delete]   [🏠 Home] │
│  ─────────────────────────────────────────────────────  │
│  SUBFORM: Recent Sales Today (last 10 records)          │
│  Product          Qty   Price    Amount    Time         │
│  ...                                                    │
└─────────────────────────────────────────────────────────┘
```

#### Key Fields

| Field | Control | Source | Notes |
|---|---|---|---|
| product_spec_id | Combo Box | tbl_product_specs JOIN tbl_products | Display: "ProductName — SpecValue" |
| quantity | Text Box | — | Default 1, numeric only |
| unit_price | Text Box | Auto from product_spec.default_selling_price | Editable |
| discount | Text Box | — | Default 0 |
| amount | Text Box | =quantity * unit_price - discount | Calculated, locked |
| payment_method_id | Combo Box | tbl_payment_methods | |
| sale_date | Text Box | =Now() | Auto, locked |
| month_id | Hidden | =Month(Now()) | Auto |
| month_name | Hidden | =Format(Now(),"mmmm") | Auto |

#### Special Behavior
- When product is selected, auto-fill `unit_price` from `default_selling_price`
- After save, if payment is "Credit/Debt" → prompt to open frm_debts to create debt record
- "Save & New" clears form and focuses on Product field
- Show running total of today's sales in footer

#### VBA Key Functions
```vba
Private Sub cboProduct_AfterUpdate()
    ' Auto-fill price from product spec
    If Not IsNull(Me.cboProduct) Then
        Me.txtUnitPrice = Nz(DLookup("default_selling_price", "tbl_product_specs", _
            "product_spec_id = " & Me.cboProduct), 0)
        Me.txtUnitPrice.SetFocus
    End If
End Sub

Private Sub txtQuantity_AfterUpdate()
    Call CalculateAmount
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    Call CalculateAmount
End Sub

Private Sub txtDiscount_AfterUpdate()
    Call CalculateAmount
End Sub

Private Sub CalculateAmount()
    Dim qty As Long, price As Currency, disc As Currency
    qty = Nz(Me.txtQuantity, 0)
    price = Nz(Me.txtUnitPrice, 0)
    disc = Nz(Me.txtDiscount, 0)
    Me.txtAmount = (qty * price) - disc
End Sub
```

---

### FORM 2: frm_purchases (Purchase Order with Details)

**Purpose:** Record stock purchases from suppliers — header + line items.  
**Tables:** tbl_purchases (header), tbl_purchase_details (lines)  
**Priority:** P0 — critical gap

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  📦 NEW PURCHASE ORDER                   [🏠] [←] [✕]  │
├─────────────────────────────────────────────────────────┤
│  PURCHASE HEADER                                        │
│  Supplier:     [▼ Select Supplier      ] [+ New]       │
│  Invoice No:   [                       ]                │
│  Date:         [ 03/05/2026            ]                │
│  Month:        [ May 2026              ] (auto)         │
├─────────────────────────────────────────────────────────┤
│  PURCHASE ITEMS (Subform — datasheet)                   │
│  ┌──────────────────┬─────┬──────────┬──────────┬────┐ │
│  │ Product/Spec     │ Qty │ Cost     │ Sell Pri │Amt │ │
│  │ [▼             ] │[  ] │ [      ] │ [      ] │  0 │ │
│  │ [+ Add Row     ] │     │          │          │    │ │
│  └──────────────────┴─────┴──────────┴──────────┴────┘ │
│  TOTAL PURCHASE VALUE:        TZS  0.00                 │
├─────────────────────────────────────────────────────────┤
│  [💾 Save Order]  [🖨️ Print GRN]  [🗑️ Delete]  [🏠]   │
└─────────────────────────────────────────────────────────┘
```

#### Subform: sub_purchase_details

| Field | Control | Notes |
|---|---|---|
| product_spec_id | Combo Box | Shows product + spec |
| quantity | Number | Updates tbl_product_specs.current_stock on save |
| unit_cost | Currency | |
| suggested_selling_price | Currency | Auto-updates default_selling_price if different |
| amount | Calculated | qty × cost |

#### Special Behavior
- On save: loop through details and **update current_stock** (`current_stock = current_stock + quantity`)
- Warn if suggested_selling_price differs from current default_selling_price by >10%
- Print button generates a Goods Received Note (GRN) report

---

### FORM 3: frm_debtors (Debtor Management)

**Purpose:** Manage credit customers — view debts, record collections.  
**Tables:** tbl_debtors, tbl_debts, tbl_debt_returns  
**Priority:** P0

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  👤 DEBTORS MANAGEMENT                  [🏠] [←] [✕]  │
├───────────────────┬─────────────────────────────────────┤
│  DEBTOR LIST      │  DEBTOR DETAILS                     │
│  [🔍 Search...  ] │  Name:    [                       ] │
│  ┌─────────────┐  │  Address: [                       ] │
│  │ John M.  ⚠️ │  │  Phone 1: [                       ] │
│  │ Mary W.  ✅ │  │  Phone 2: [                       ] │
│  │ Peter S. ⚠️ │  │  NIDA ID: [                       ] │
│  │ [+ New   ] │  │                                     │
│  └─────────────┘  │  BALANCE:  TZS  0.00               │
│                   │  [📞 Call] [📋 View All Debts]      │
├───────────────────┴─────────────────────────────────────┤
│  DEBTS (Subform for selected debtor)                    │
│  Date       Product       Qty  Amount  Due Date  Status │
│  ...                                                    │
├─────────────────────────────────────────────────────────┤
│  RECORD PAYMENT                                         │
│  Debt Item: [▼ Select Debt         ]                    │
│  Amount:    [          ]  Method: [▼ Cash / Mobile ]   │
│  Date:      [ Today    ]  Comment:[                ]    │
│  [💾 Record Payment]                                    │
└─────────────────────────────────────────────────────────┘
```

#### Key Logic
- Balance = SUM(tbl_debts.amount) - SUM(tbl_debt_returns.amount) for this debtor
- ⚠️ icon in list = debtor has overdue debt (expected_payment_date < Today())
- "Record Payment" section writes to tbl_debt_returns

---

### FORM 4: frm_expenses_obligations (Expense Bills & Payments)

**Purpose:** View upcoming/overdue payment obligations and record payments.  
**Tables:** tbl_payment_obligations, tbl_payments, tbl_expense_items  
**Priority:** P1

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  💳 EXPENSES & BILLS                    [🏠] [←] [✕]  │
├─────────────────────────────────────────────────────────┤
│  FILTER: [▼ All Types] [▼ All Statuses] [📅 Date Range]│
├─────────────────────────────────────────────────────────┤
│  OBLIGATIONS LIST                                       │
│  Description      Type    Due Date   Owed    Status    │
│  Rent — May       Rent    01/05/2026 150,000 ⚠️ Overdue│
│  Electricity      Utility 15/05/2026  45,000 🔵 Pending│
│  [+ New Obligation]                                    │
├─────────────────────────────────────────────────────────┤
│  RECORD PAYMENT (for selected obligation)               │
│  Obligation: [ Auto-filled from selection             ] │
│  Pay Amount: [          ]  Date: [ Today  ]            │
│  Method:     [▼ Cash / Mobile / Bank  ]                │
│  Reference:  [                        ]                │
│  [💾 Pay Now]  [💳 Prepayment]                         │
└─────────────────────────────────────────────────────────┘
```

---

### FORM 5: frm_assets (Asset Register)

**Purpose:** Manage business fixed assets.  
**Tables:** tbl_assets, tbl_asset_types, tbl_asset_categories  
**Priority:** P1

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  🏭 ASSET REGISTER                      [🏠] [←] [✕]  │
├─────────────────────────────────────────────────────────┤
│  ASSET LIST (left panel)    │  ASSET DETAILS (right)    │
│  [🔍 Filter by category]    │  Name:     [            ] │
│  ┌──────────────────────┐   │  Type:     [▼           ] │
│  │ 📱 Samsung TV        │   │  Category: [▼           ] │
│  │ 💻 Laptop HP         │   │  Value:    TZS [        ] │
│  │ 🚗 Delivery Motorbike│   │  Checked:  [  Date      ] │
│  │ [+ New Asset    ]    │   │  Notes:    [            ] │
│  └──────────────────────┘   │                           │
│                             │  TOTAL VALUE:  TZS 0.00   │
├─────────────────────────────┴───────────────────────────┤
│  [➕ New]  [💾 Save]  [🗑️ Delete]  [🖨️ Print Register] │
└─────────────────────────────────────────────────────────┘
```

---

### FORM 6: frm_liabilities (Loans & Payables)

**Purpose:** Track business loans and manage repayment schedules.  
**Tables:** tbl_liability_items, tbl_liability_types, tbl_liability_payment_details  
**Priority:** P1

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  🏦 LIABILITIES & LOANS                 [🏠] [←] [✕]  │
├─────────────────────────────────────────────────────────┤
│  Liability Name:   [                              ]     │
│  Type:             [▼ Loan / Supplier Payable... ]      │
│  Original Amount:  TZS [                         ]      │
│  Current Balance:  TZS [ Auto-calculated         ]      │
│  Start Date:       [          ]  Rate: [     ] %        │
│  Maturity Date:    [          ]  Per Payment: TZS [   ] │
│  Active:           [✓]                                  │
├─────────────────────────────────────────────────────────┤
│  PAYMENT HISTORY (Subform)                              │
│  Date         Principal    Interest    Balance After    │
│  ...                                                    │
├─────────────────────────────────────────────────────────┤
│  RECORD REPAYMENT                                       │
│  Principal:  TZS [       ]  Interest: TZS [       ]    │
│  Date:       [  Today    ]  Method:   [▼            ]  │
│  [💾 Record Repayment]                                  │
└─────────────────────────────────────────────────────────┘
```

---

### FORM 7: frm_drawings (Owner Withdrawals)

**Purpose:** Record owner's personal withdrawals from the business.  
**Tables:** tbl_drawings, tbl_drawing_categories, tbl_product_specs  
**Priority:** P1

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  💸 OWNER DRAWINGS                      [🏠] [←] [✕]  │
├─────────────────────────────────────────────────────────┤
│  DRAWINGS LIST (Subform — datasheet)                    │
│  Date    Category    Product/Spec    Qty   Amount  Notes│
│  ...                                                    │
├─────────────────────────────────────────────────────────┤
│  NEW DRAWING                                            │
│  Category:   [▼ Personal Use / Goods / Cash    ]       │
│  Product:    [▼ Select Product (if goods)      ]       │
│  Quantity:   [    ]   Unit Price: TZS [         ]      │
│  Amount:     TZS [ Auto ]                              │
│  Date:       [ Today ]   Notes: [               ]      │
├─────────────────────────────────────────────────────────┤
│  MONTH TOTAL: TZS 0.00   YEAR TOTAL: TZS 0.00          │
│  [💾 Save]   [🗑️ Delete]   [🖨️ Monthly Report]        │
└─────────────────────────────────────────────────────────┘
```

---

### FORM 8: frm_returns (Return Inwards & Outwards)

**Purpose:** Record goods returned by customers (inward) or to suppliers (outward).  
**Tables:** tbl_return_inwards, tbl_return_outwards, tbl_sales, tbl_purchase_details  
**Priority:** P1 — tbl_return_inwards has 0 records

#### Layout (Tabbed)

```
┌─────────────────────────────────────────────────────────┐
│  📦 RETURNS                             [🏠] [←] [✕]  │
│  [📥 RETURN INWARDS] | [📤 RETURN OUTWARDS]            │
├─────────────────────────────────────────────────────────┤
│  TAB 1: RETURN INWARDS (Customer → Shop)               │
│  Original Sale: [▼ Search by date/product    ]         │
│  Product shown: [ Auto-filled                ]         │
│  Return Qty:    [    ]   Unit Price: TZS [   ]         │
│  Reason:        [▼ Defective / Wrong item / Other ]    │
│  Date:          [ Today ]                              │
│  AMOUNT:        TZS [ Auto ]                           │
│  [💾 Save Return]  — Stock auto-incremented            │
├─────────────────────────────────────────────────────────┤
│  TAB 2: RETURN OUTWARDS (Shop → Supplier)              │
│  Original Purchase Detail: [▼ Search      ]            │
│  Return Qty:    [    ]   Unit Price: TZS [ ]           │
│  Reason:        [▼                        ]            │
│  Date:          [ Today ]                              │
│  AMOUNT:        TZS [ Auto ]                           │
│  [💾 Save Return]  — Stock auto-decremented            │
└─────────────────────────────────────────────────────────┘
```

---

### FORM 9: frm_suppliers (Supplier Management)

**Purpose:** Add/edit suppliers; view purchase history per supplier.  
**Tables:** tbl_suppliers, tbl_purchases, tbl_purchase_details  
**Priority:** P2

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  🏭 SUPPLIERS                           [🏠] [←] [✕]  │
├───────────────────────┬─────────────────────────────────┤
│  SUPPLIER LIST        │  DETAILS                        │
│  [🔍 Search...  ]     │  Name:    [                   ] │
│  ┌─────────────────┐  │  Contact: [                   ] │
│  │ Supplier A      │  │  Phone:   [                   ] │
│  │ Supplier B      │  │  Address: [                   ] │
│  │ [+ New      ]   │  │                                 │
│  └─────────────────┘  │  TOTAL PURCHASES: TZS 0.00     │
│                       │  LAST PURCHASE:   [date]        │
├───────────────────────┴─────────────────────────────────┤
│  PURCHASE HISTORY (Subform)                             │
│  Date        Invoice     Items   Total Value            │
│  ...                                                    │
└─────────────────────────────────────────────────────────┘
```

---

### FORM 10: frm_reports_hub (Report Launcher)

**Purpose:** Central hub to open all reports with date filter parameters.  
**Priority:** P1 — needed alongside each report

#### Layout

```
┌─────────────────────────────────────────────────────────┐
│  🖨️ REPORTS                             [🏠] [←] [✕]  │
├─────────────────────────────────────────────────────────┤
│  DATE FILTER (applies to all reports)                   │
│  From: [ 01/05/2026 ]  To: [ 31/05/2026 ]  [This Month]│
│        [Today] [This Week] [This Month] [This Year] [Custom]│
├─────────────────────────────────────────────────────────┤
│  SALES REPORTS          STOCK REPORTS                   │
│  [📊 Daily Sales]       [📦 Current Stock Levels]      │
│  [📈 Monthly Summary]   [⚠️  Low Stock Alert]          │
│  [🏆 Top Products]      [📋 Purchase History]          │
│  [👤 Sales by Method]                                   │
│                                                         │
│  FINANCE REPORTS        DEBTOR REPORTS                  │
│  [📉 P&L Statement]     [👥 Debtor Listing]            │
│  [💰 Cash Flow]         [⏰ Debtor Ageing]              │
│  [⚖️  Balance Sheet]    [💸 Debt Collections]          │
│  [💳 Expense Summary]                                   │
└─────────────────────────────────────────────────────────┘
```

---

## 6. Reports — Complete Specification

---

### REPORT 1: rpt_sales_daily (Daily Sales Summary)

**Purpose:** End-of-day cash reconciliation summary.  
**Data Source:** tbl_sales JOIN tbl_product_specs JOIN tbl_products JOIN tbl_payment_methods  
**Parameters:** Date (default = today)  
**Priority:** P0 — most critical report

#### Layout

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
                   KIYABO DUKA
              DAILY SALES SUMMARY
              Date: Saturday, 3 May 2026
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SALES BY PAYMENT METHOD
──────────────────────────────────────────────
Payment Method     Transactions    Amount (TZS)
Cash                      45         234,500.00
Mobile Money              12          89,000.00
Card                       3          45,000.00
Credit (Debt)              2          30,000.00
──────────────────────────────────────────────
TOTAL                     62         398,500.00

TOP 10 PRODUCTS SOLD TODAY
──────────────────────────────────────────────
Product                    Qty    Revenue (TZS)
Product A — Spec X          15        45,000.00
...
──────────────────────────────────────────────

SUMMARY
  Gross Sales:          TZS   398,500.00
  Discounts Given:      TZS    (2,500.00)
  NET SALES:            TZS   396,000.00

  Printed: 03/05/2026 18:30  [Operator: Admin]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

#### Query (qry_sales_daily)

```sql
SELECT 
    pm.payment_method,
    COUNT(s.sale_id) AS total_transactions,
    SUM(s.amount) AS total_amount,
    SUM(s.discount) AS total_discount
FROM tbl_sales s
LEFT JOIN tbl_payment_methods pm ON s.payment_method_id = pm.payment_method_id
WHERE s.sale_date >= Date() AND s.sale_date < Date() + 1
GROUP BY pm.payment_method
ORDER BY total_amount DESC;
```

---

### REPORT 2: rpt_stock_levels (Stock Level Report)

**Purpose:** Full inventory snapshot with low-stock alerts.  
**Data Source:** tbl_product_specs JOIN tbl_products JOIN tbl_types JOIN tbl_categories  
**Priority:** P0

#### Layout

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
                      KIYABO DUKA
                 STOCK LEVELS REPORT
                 As at: 3 May 2026
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚠️  LOW STOCK ITEMS (Below Reorder Level)
────────────────────────────────────────────────────────
Product              Spec      Stock  Reorder  Shortage
Product A            Large         2        5       -3 ⚠️
...
────────────────────────────────────────────────────────

FULL STOCK LISTING (Grouped by Category)
CATEGORY: Electronics
  TYPE: Phones
────────────────────────────────────────────────────────
Product          Spec       Cost Price  Stock  Stock Value
...
────────────────────────────────────────────────────────
Category Subtotal:                      ---    TZS 0.00

GRAND TOTAL STOCK VALUE:               TZS  0,000,000.00
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

#### Query (qry_stock_levels)

```sql
SELECT 
    c.category_name,
    t.type_name,
    p.product_name,
    sv.spec_value,
    ps.default_cost_price,
    ps.default_selling_price,
    ps.current_stock,
    ps.reorder_level,
    ps.current_stock - ps.reorder_level AS stock_gap,
    ps.current_stock * ps.default_cost_price AS stock_value,
    IIf(ps.current_stock <= ps.reorder_level, "⚠️ LOW", "OK") AS stock_status
FROM (((tbl_product_specs ps
INNER JOIN tbl_products p ON ps.product_id = p.product_id)
INNER JOIN tbl_spec_values sv ON ps.spec_value_id = sv.spec_value_id)
INNER JOIN tbl_types t ON p.type_id = t.type_id)
INNER JOIN tbl_categories c ON t.category_id = c.category_id
ORDER BY c.category_name, t.type_name, p.product_name;
```

---

### REPORT 3: rpt_debtor_ageing (Debtor Ageing Report)

**Purpose:** Show who owes what, for how long, and what is overdue.  
**Data Source:** tbl_debtors, tbl_debts, tbl_debt_returns  
**Priority:** P0

#### Layout

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
                       KIYABO DUKA
                  DEBTOR AGEING REPORT
                  As at: 3 May 2026
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Debtor      Phone         Current  1-30 days  31-60d  60+d  Total
John Mwamba 0754-xxx-xxx       0   50,000     30,000    0   80,000
Mary Wangui 0712-xxx-xxx   5,000        0          0    0    5,000
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TOTAL                      5,000   50,000     30,000    0   85,000
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

#### Ageing Buckets (based on expected_payment_date)

```sql
-- Ageing Query
SELECT 
    d.debtor_name,
    d.phone_number_1,
    SUM(IIf(DateDiff("d", dt.expected_payment_date, Date()) <= 0, 
        dt.amount - Nz(paid.total_paid,0), 0)) AS current_balance,
    SUM(IIf(DateDiff("d", dt.expected_payment_date, Date()) BETWEEN 1 AND 30, 
        dt.amount - Nz(paid.total_paid,0), 0)) AS overdue_30,
    SUM(IIf(DateDiff("d", dt.expected_payment_date, Date()) BETWEEN 31 AND 60, 
        dt.amount - Nz(paid.total_paid,0), 0)) AS overdue_60,
    SUM(IIf(DateDiff("d", dt.expected_payment_date, Date()) > 60, 
        dt.amount - Nz(paid.total_paid,0), 0)) AS overdue_over_60
FROM tbl_debtors d
INNER JOIN tbl_debts dt ON d.debtor_id = dt.debtor_id
LEFT JOIN (
    SELECT debt_id, SUM(amount) AS total_paid FROM tbl_debt_returns GROUP BY debt_id
) AS paid ON dt.sale_id = paid.debt_id
GROUP BY d.debtor_id, d.debtor_name, d.phone_number_1
HAVING SUM(dt.amount - Nz(paid.total_paid,0)) > 0
ORDER BY overdue_over_60 DESC;
```

---

### REPORT 4: rpt_profit_loss (Monthly P&L Statement)

**Purpose:** Shows net profit/loss for a given period.  
**Data Source:** tbl_sales, tbl_purchase_details, tbl_payments, tbl_debt_returns  
**Priority:** P1

#### Layout

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
               KIYABO DUKA
        PROFIT & LOSS STATEMENT
        Month: May 2026
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

REVENUE
  Sales (Cash)             TZS   800,000.00
  Sales (Mobile)           TZS   200,000.00
  Debt Collections         TZS    50,000.00
  ─────────────────────────────────────────
  GROSS REVENUE            TZS 1,050,000.00

COST OF GOODS SOLD
  Opening Stock Value      TZS   300,000.00
  + Purchases              TZS   400,000.00
  - Closing Stock Value    TZS  (250,000.00)
  ─────────────────────────────────────────
  COST OF GOODS SOLD       TZS   450,000.00
  ─────────────────────────────────────────
  GROSS PROFIT             TZS   600,000.00

OPERATING EXPENSES
  Rent                     TZS   150,000.00
  Electricity              TZS    45,000.00
  Salaries                 TZS   100,000.00
  Other                    TZS    30,000.00
  ─────────────────────────────────────────
  TOTAL EXPENSES           TZS   325,000.00

  ═════════════════════════════════════════
  NET PROFIT               TZS   275,000.00
  ═════════════════════════════════════════

  Owner Drawings           TZS   (50,000.00)
  ─────────────────────────────────────────
  RETAINED EARNINGS        TZS   225,000.00
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

### REPORT 5: rpt_sales_monthly (Monthly Sales Trend)

**Purpose:** Month-by-month sales summary for the year.  
**Priority:** P1

```sql
SELECT 
    month_name,
    month_id,
    COUNT(sale_id) AS transactions,
    SUM(amount) AS gross_sales,
    SUM(discount) AS total_discounts,
    SUM(amount) - SUM(discount) AS net_sales
FROM tbl_sales
WHERE sale_date >= DateSerial(Year(Date()), 1, 1)
GROUP BY month_id, month_name
ORDER BY month_id;
```

---

### REPORT 6: rpt_expense_summary (Expense Summary)

**Purpose:** Shows expenses by type for a period — paid vs outstanding.  
**Priority:** P1

```sql
SELECT 
    et.expense_type,
    ei.item_name,
    po.obligation_date,
    po.due_date,
    po.amount_due,
    po.amount_paid,
    po.balance,
    po.payment_status
FROM tbl_payment_obligations po
LEFT JOIN tbl_expense_items ei ON po.expense_item_id = ei.expense_item_id
LEFT JOIN tbl_expense_types et ON ei.expense_type_id = et.expense_type_id
ORDER BY et.expense_type, po.due_date;
```

---

### REPORT 7: rpt_purchase_history (Purchase History by Supplier)

**Purpose:** What was bought, from whom, at what cost.  
**Priority:** P2

---

### REPORT 8: rpt_cash_flow (Cash Flow Statement)

**Purpose:** Money in vs money out for a period.  
**Priority:** P2

#### Structure

```
CASH INFLOWS
  Cash Sales
  Mobile Money Sales
  Debt Collections
  Total Inflows

CASH OUTFLOWS
  Purchases (stock)
  Rent Paid
  Utilities Paid
  Salaries Paid
  Loan Repayments
  Owner Drawings
  Total Outflows

NET CASH MOVEMENT
Opening Cash Balance
CLOSING CASH BALANCE
```

---

### REPORT 9: rpt_balance_sheet (Balance Sheet)

**Purpose:** Financial snapshot — assets vs liabilities.  
**Priority:** P2

---

### REPORT 10: rpt_top_products (Top Products by Sales)

**Purpose:** Best-selling products ranked by revenue.  
**Priority:** P2

```sql
SELECT TOP 20
    p.product_name,
    sv.spec_value,
    SUM(s.quantity) AS total_qty_sold,
    SUM(s.amount) AS total_revenue,
    AVG(s.unit_price) AS avg_selling_price
FROM tbl_sales s
INNER JOIN tbl_product_specs ps ON s.product_spec_id = ps.product_spec_id
INNER JOIN tbl_products p ON ps.product_id = p.product_id
INNER JOIN tbl_spec_values sv ON ps.spec_value_id = sv.spec_value_id
GROUP BY p.product_name, sv.spec_value
ORDER BY total_revenue DESC;
```

---

## 7. VBA Module Reference

### modHelpers — Utility Functions

```vba
' ============================================================
' UTILITY FUNCTIONS — modHelpers
' ============================================================

' Open a form safely (check if already open)
Public Sub OpenFormSafe(strFormName As String, Optional varWhere As Variant)
    If Not IsFormOpen(strFormName) Then
        If IsMissing(varWhere) Then
            DoCmd.OpenForm strFormName
        Else
            DoCmd.OpenForm strFormName, , , varWhere
        End If
    Else
        Forms(strFormName).SetFocus
    End If
End Sub

' Check if a form is currently open
Public Function IsFormOpen(strFormName As String) As Boolean
    Dim frm As AccessObject
    For Each frm In CurrentProject.AllForms
        If frm.Name = strFormName And frm.IsLoaded Then
            IsFormOpen = True
            Exit Function
        End If
    Next frm
    IsFormOpen = False
End Function

' Format currency in TZS
Public Function FormatTZS(dblAmount As Double) As String
    FormatTZS = "TZS " & Format(dblAmount, "#,##0.00")
End Function

' Confirm delete action
Public Function ConfirmDelete(strItemName As String) As Boolean
    ConfirmDelete = (MsgBox("Are you sure you want to delete '" & strItemName & "'?" & vbCrLf & _
        "This cannot be undone.", vbYesNo + vbCritical, "Confirm Delete") = vbYes)
End Function

' Get current month ID
Public Function GetCurrentMonthID() As Integer
    GetCurrentMonthID = Month(Date)
End Function

' Get current month name
Public Function GetCurrentMonthName() As String
    GetCurrentMonthName = Format(Date, "mmmm")
End Function

' Navigate to home (frm_master)
Public Sub GoHome()
    OpenFormSafe "frm_master"
End Sub
```

### modStockCRUD — Stock Update Functions

```vba
' ============================================================
' STOCK MANAGEMENT — modStockCRUD
' ============================================================

' Increase stock after purchase
Public Sub IncreaseStock(lngProductSpecID As Long, lngQty As Long)
    Dim strSQL As String
    strSQL = "UPDATE tbl_product_specs SET current_stock = current_stock + " & lngQty & _
             " WHERE product_spec_id = " & lngProductSpecID
    CurrentDb.Execute strSQL, dbFailOnError
End Sub

' Decrease stock after sale (with check)
Public Function DecreaseStock(lngProductSpecID As Long, lngQty As Long) As Boolean
    Dim lngCurrentStock As Long
    lngCurrentStock = Nz(DLookup("current_stock", "tbl_product_specs", _
        "product_spec_id = " & lngProductSpecID), 0)
    
    If lngCurrentStock < lngQty Then
        MsgBox "Insufficient stock! Available: " & lngCurrentStock & " " & _
               "Requested: " & lngQty, vbExclamation, "Stock Alert"
        DecreaseStock = False
        Exit Function
    End If
    
    Dim strSQL As String
    strSQL = "UPDATE tbl_product_specs SET current_stock = current_stock - " & lngQty & _
             " WHERE product_spec_id = " & lngProductSpecID
    CurrentDb.Execute strSQL, dbFailOnError
    
    ' Check if now at reorder level
    Dim lngNewStock As Long
    lngNewStock = Nz(DLookup("current_stock", "tbl_product_specs", _
        "product_spec_id = " & lngProductSpecID), 0)
    Dim lngReorder As Long
    lngReorder = Nz(DLookup("reorder_level", "tbl_product_specs", _
        "product_spec_id = " & lngProductSpecID), 5)
    
    If lngNewStock <= lngReorder Then
        MsgBox "⚠️ Low Stock Alert!" & vbCrLf & "Stock is now " & lngNewStock & _
               " (Reorder level: " & lngReorder & ")", vbInformation, "Low Stock"
    End If
    
    DecreaseStock = True
End Function

' Adjust stock for return inward (customer returns goods)
Public Sub ReturnInwardStockAdjust(lngProductSpecID As Long, lngQty As Long)
    Call IncreaseStock(lngProductSpecID, lngQty)
End Sub

' Adjust stock for return outward (shop returns to supplier)
Public Sub ReturnOutwardStockAdjust(lngProductSpecID As Long, lngQty As Long)
    Call DecreaseStock lngProductSpecID, lngQty
End Sub
```

---

## 8. Schema Improvements

### 8.1 Issues to Fix

| Issue | Table | Fix |
|---|---|---|
| PK named `sale_id` in tbl_debts | tbl_debts | Rename to `debt_id` (update all FK references) |
| No receipt grouping for sales | tbl_sales | Add `receipt_id` field linking multiple items to one receipt |
| No barcode/SKU field | tbl_products | Add `barcode` Text(50) field |
| month_name stored redundantly | Multiple | Consider removing — derive from month_id |
| Required=No on logical required fields | Multiple | Set Required=Yes on name/date fields |
| No audit trail | — | Add `created_by`, `modified_by`, `modified_date` to key tables |

### 8.2 Suggested New Field: receipt_id in tbl_sales

```sql
ALTER TABLE tbl_sales ADD COLUMN receipt_id Long;
```

Then generate receipt numbers like `RCP-20260503-001` using VBA:

```vba
Public Function GetNextReceiptID() As String
    Dim strDate As String
    Dim lngCount As Long
    strDate = Format(Date, "yyyymmdd")
    lngCount = Nz(DCount("*", "tbl_sales", "receipt_id LIKE 'RCP-" & strDate & "*'"), 0) + 1
    GetNextReceiptID = "RCP-" & strDate & "-" & Format(lngCount, "000")
End Function
```

### 8.3 Suggested New Field: barcode in tbl_products

```sql
ALTER TABLE tbl_products ADD COLUMN barcode Text(50);
```

Then in frm_sales_pos, add a barcode scanner input field:

```vba
Private Sub txtBarcode_AfterUpdate()
    ' Find product by barcode and auto-select
    Dim lngSpecID As Long
    lngSpecID = Nz(DLookup("product_spec_id", "tbl_product_specs ps " & _
        "INNER JOIN tbl_products p ON ps.product_id = p.product_id", _
        "p.barcode = '" & Me.txtBarcode & "'"), 0)
    If lngSpecID > 0 Then
        Me.cboProduct = lngSpecID
        Call cboProduct_AfterUpdate
    Else
        MsgBox "Product not found for barcode: " & Me.txtBarcode, vbExclamation
    End If
    Me.txtBarcode = ""  ' Clear for next scan
End Sub
```

---

## 9. Build Roadmap & Priority Order

### Phase 1 — Critical (Do These Now)

| # | Item | Type | Est. Time |
|---|---|---|---|
| 1 | frm_master (dashboard) | Form | 3–4 hours |
| 2 | frm_purchases (header + detail subform) | Form | 2–3 hours |
| 3 | frm_debtors (debtor + debt entry + collections) | Form | 3–4 hours |
| 4 | rpt_sales_daily | Report | 1–2 hours |
| 5 | rpt_stock_levels + low stock alert | Report | 1–2 hours |
| 6 | rpt_debtor_ageing | Report | 2 hours |

**Phase 1 Total: ~14–17 hours of focused work**

### Phase 2 — Important (Do Next Week)

| # | Item | Type | Est. Time |
|---|---|---|---|
| 7 | frm_expenses_obligations | Form | 2–3 hours |
| 8 | frm_returns (inward + outward tabs) | Form | 2 hours |
| 9 | rpt_profit_loss | Report | 2–3 hours |
| 10 | frm_drawings (main + single) | Form | 1–2 hours |
| 11 | rpt_expense_summary | Report | 1 hour |
| 12 | frm_reports_hub | Form | 1 hour |

### Phase 3 — Complete the Picture

| # | Item | Type | Est. Time |
|---|---|---|---|
| 13 | frm_assets | Form | 1–2 hours |
| 14 | frm_liabilities | Form | 2 hours |
| 15 | frm_suppliers | Form | 1 hour |
| 16 | rpt_sales_monthly | Report | 1 hour |
| 17 | rpt_cash_flow | Report | 2 hours |
| 18 | rpt_balance_sheet | Report | 2 hours |
| 19 | rpt_top_products | Report | 1 hour |
| 20 | Add receipt_id and barcode fields | Schema | 1 hour |

---

## 10. Speed-Up Tools & Boilerplates

This section is the most important for your productivity. Use these tools to 10× your development speed.

### 10.1 The Fastest Way to Build Forms in Access

**Method A: Wizard + Customize (Recommended for 80% of forms)**

1. In the ribbon → `Create` → `Form Wizard`
2. Select your table/query as record source
3. Choose fields you need
4. Select "Columnar" layout
5. Click Finish
6. Switch to **Design View**
7. Add your custom header, footer, and navigation buttons
8. Set combo boxes for FK fields using the Combo Box Wizard

**Time saving: Form wizard generates 60% of the work in 2 minutes.**

**Method B: Build from a Query (Best for complex forms)**

1. First build your query (qry_sales_form etc.) with all joins
2. `Create` → `Form Design` → set Record Source to your query
3. Drag fields from the Field List panel to the form
4. This gives you full control from the start

**Method C: Copy an Existing Form (Fastest for similar forms)**

1. Right-click an existing form in Navigation Pane → Copy → Paste → new name
2. Modify the Record Source and adjust fields
3. Keep the header/footer layout — just change the title

**Use this for:** frm_debtors is very similar to frm_suppliers — copy one, modify the other.

---

### 10.2 Access Form Builder Tools (Online & Desktop)

#### Built-in Access Supercharger: The Property Sheet

The single most powerful Access feature most developers underuse:

- **Format** tab: Set `Format` for currency fields → `Currency` or `#,##0.00`
- **Data** tab: Set `Input Mask` for phone numbers → `000\-000\-000;0;_`
- **Event** tab: This is where all VBA goes — double-click any event to create the handler
- **Other** tab: Set `Tab Stop = No` on calculated fields, `Auto Tab = Yes` on short fields

#### Speed Trick: Field List Drag

In form Design View → `Design` ribbon → `Add Existing Fields`.
Drag multiple fields at once by holding Ctrl+Click, then drag them all together.
This places all fields with their labels in one action.

---

### 10.3 ChatGPT / Claude for VBA Code Generation

**You are already using Claude!** Here's exactly how to use it to generate 80% of your VBA code:

**Prompt Template for Forms:**

```
I'm building a Microsoft Access 2016+ form called [frm_name].

Tables involved:
- tbl_sales (sale_id, product_spec_id, quantity, unit_price, discount, amount, sale_date, payment_method_id)
- tbl_product_specs (product_spec_id, product_id, default_selling_price, current_stock)

I need:
1. VBA for the On Load event to initialize the form
2. VBA for the product combo box AfterUpdate to auto-fill price
3. VBA to calculate amount = (quantity × unit_price) - discount
4. VBA for the Save button that validates required fields first
5. VBA to update current_stock after saving a sale

Write clean, commented VBA code for all of these.
```

**Prompt Template for Reports:**

```
I need an Access SQL query for a report showing:
- All sales for [date range]
- Grouped by payment method
- Showing count of transactions and sum of amounts
- Also show the top 10 products

Tables: tbl_sales, tbl_payment_methods, tbl_products, tbl_product_specs, tbl_spec_values

Write the SQL query I can use as the Record Source for an Access report.
```

**Use Claude throughout your development.** Paste your schema and ask for specific queries, VBA procedures, or form logic.

---

### 10.4 Microsoft Access Templates

**Free professional Access templates you can use as starting points:**

1. **Microsoft Office Template Gallery**  
   `office.com/templates` → search "Access" → "Inventory" or "Sales"  
   Download the Northwind database — it's the gold standard template

2. **Northwind 2.0 (Updated)**  
   Search: "Access Northwind 2.0 download"  
   Has: Products, Orders, Customers, Suppliers — very close to your schema  
   Use it to copy form design patterns

3. **Access Asset Tracking Template**  
   `office.com/templates` → "Access Asset Tracking"  
   Copy the asset form design for your frm_assets

---

### 10.5 Useful Access Add-Ins & Tools

| Tool | Purpose | Cost | Where |
|---|---|---|---|
| **Total Access Components** | Professional controls (progress bars, calendars) | Paid | FMS Inc |
| **MZ-Tools** | Code cleanup, module navigator, find/replace | Free/Paid | mztools.com |
| **Access Speed Optimization Tool** | Find slow queries, unused objects | Free | fmsinc.com |
| **Database Documenter** (built-in) | Access → Database Tools → Database Documenter — generates schema report | Free | Built-in |

---

### 10.6 Query Builder Productivity Tips

**Never write SQL from scratch in Access.** Use these steps:

1. `Create` → `Query Design`
2. Add your tables visually (drag to create joins)
3. Add fields by double-clicking or dragging
4. Use the criteria row for WHERE clauses
5. Then click `View` → `SQL View` to see the generated SQL — copy and refine it

**Best Access query tricks:**

```sql
-- Date ranges using parameters (prompts user)
WHERE sale_date BETWEEN [Enter Start Date:] AND [Enter End Date:]

-- Current month
WHERE Month(sale_date) = Month(Date()) AND Year(sale_date) = Year(Date())

-- Current year
WHERE Year(sale_date) = Year(Date())

-- Null-safe amount calculations
IIf(IsNull([discount]), 0, [discount])
-- or more cleanly:
Nz([discount], 0)
```

---

### 10.7 Report Design Speed Tips

**The 4-step report method:**

1. **Build the query first** — get the data right before touching the report
2. **Use Report Wizard** → `Create` → `Report Wizard` → select your query
3. Add grouping levels (by category, by month, etc.) in the wizard
4. Switch to **Design View** to add totals, headers, and formatting

**Professional report tricks:**

```vba
' In Report Header — show parameter values
Me.lblReportPeriod.Caption = "Period: " & Format([Forms]![frm_reports_hub]![dtFrom], "d mmm yyyy") & _
    " to " & Format([Forms]![frm_reports_hub]![dtTo], "d mmm yyyy")

' Alternating row colors (in Detail section — On Format event)
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If Me.CurrentRecord Mod 2 = 0 Then
        Me.Detail.BackColor = RGB(245, 245, 245)
    Else
        Me.Detail.BackColor = RGB(255, 255, 255)
    End If
End Sub

' Running total (use text box with Control Source)
' Set Running Sum property to "Over All" or "Over Group"
' =Sum([amount])  → in group footer
```

---

## 11. Access VBA Boilerplate Library

Copy-paste these into your modules. They will save you hours.

### Standard Form Template

Every new form should start with this code structure:

```vba
Option Compare Database
Option Explicit

' ============================================================
' FORM: frm_[name]
' Purpose: [describe purpose]
' Tables: [list tables used]
' Author: [you]
' Created: [date]
' Modified: [date] — [what changed]
' ============================================================

Private Sub Form_Load()
    ' Set form title
    Me.Caption = "[Form Title]"
    
    ' Populate combo boxes if not bound to queries
    ' (Usually handled by Row Source property — set in design view)
    
    ' Set default values
    If Me.NewRecord Then
        Me.txtDate = Now()
        Me.txtMonthID = Month(Now())
        Me.txtMonthName = Format(Now(), "mmmm")
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Validation before saving
    If IsNull(Me.cboRequiredField) Or Me.cboRequiredField = "" Then
        MsgBox "Please select a [required field] before saving.", vbExclamation, "Required Field"
        Me.cboRequiredField.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If IsNull(Me.txtAmount) Or Me.txtAmount <= 0 Then
        MsgBox "Amount must be greater than zero.", vbExclamation, "Invalid Amount"
        Me.txtAmount.SetFocus
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Form_AfterUpdate()
    ' After successful save
    MsgBox "Record saved successfully.", vbInformation, "Saved"
End Sub

Private Sub btnSave_Click()
    If Me.Dirty Then
        Me.Dirty = False  ' Triggers BeforeUpdate → AfterUpdate
    Else
        MsgBox "No changes to save.", vbInformation, "Save"
    End If
End Sub

Private Sub btnNew_Click()
    DoCmd.GoToRecord , , acNewRec
    Me.cboFirstField.SetFocus
End Sub

Private Sub btnDelete_Click()
    If ConfirmDelete(Me.txtRecordName) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnHome_Click()
    Call GoHome
End Sub

Private Sub btnPrint_Click()
    ' Open related report filtered to this record
    DoCmd.OpenReport "rpt_[related_report]", acViewPreview, , _
        "[primary_key] = " & Me.txtPrimaryKey
End Sub
```

### Standard Report Parameters Form

```vba
' For frm_reports_hub or any report launcher

Private Sub btnThisMonth_Click()
    Me.dtFrom = DateSerial(Year(Date), Month(Date), 1)
    Me.dtTo = DateSerial(Year(Date), Month(Date) + 1, 0)
End Sub

Private Sub btnToday_Click()
    Me.dtFrom = Date
    Me.dtTo = Date
End Sub

Private Sub btnThisYear_Click()
    Me.dtFrom = DateSerial(Year(Date), 1, 1)
    Me.dtTo = DateSerial(Year(Date), 12, 31)
End Sub

Private Sub btnThisWeek_Click()
    Me.dtFrom = Date - Weekday(Date, vbMonday) + 1
    Me.dtTo = Me.dtFrom + 6
End Sub

' Open report with date filter
Public Function OpenDateFilteredReport(strReportName As String) As Boolean
    If IsNull(Me.dtFrom) Or IsNull(Me.dtTo) Then
        MsgBox "Please select a date range.", vbExclamation
        Exit Function
    End If
    
    Dim strFilter As String
    strFilter = "sale_date BETWEEN #" & Format(Me.dtFrom, "mm/dd/yyyy") & _
                "# AND #" & Format(Me.dtTo + 1, "mm/dd/yyyy") & "#"
    
    DoCmd.OpenReport strReportName, acViewPreview, , strFilter
    OpenDateFilteredReport = True
End Function
```

---

## 12. Online Tools & Resources

### 12.1 Essential Bookmarks

| Resource | URL | What For |
|---|---|---|
| Access documentation | docs.microsoft.com/office/access | Official reference |
| Access World Forums | accessworld.net | Best Access community forum |
| UtterAccess | utteraccess.com | #1 Access help forum |
| Stack Overflow — ms-access tag | stackoverflow.com/questions/tagged/ms-access | Code Q&A |
| Allen Browne's Access tips | allenbrowne.com | Advanced techniques, free code |
| Dev Ashish's Access tips | trigeminal.com | More free VBA examples |

### 12.2 SQL Tools

| Tool | Purpose |
|---|---|
| **DB Browser for SQLite** | Test SQL queries without affecting your .accdb |
| **SQL Fiddle** (sqlfiddle.com) | Test SQL online |
| **AI Query Builder** | Ask Claude: "Write an Access SQL query that..." |

### 12.3 The 20-Minute Form Method

When you sit down to build any new form, follow this exact sequence:

1. **(2 min)** Build or identify the query → save it as `qry_[formname]`
2. **(1 min)** `Create` → `Form Wizard` → select your query → Columnar → Finish
3. **(3 min)** Delete the auto-generated header, add your standard header template (copy from another form)
4. **(5 min)** Change Text Boxes to Combo Boxes for all FK fields. Set Row Source to lookup table.
5. **(3 min)** Add footer buttons: New, Save, Delete, Print, Close, Home
6. **(3 min)** Open VBA editor (Alt+F11) → paste your standard Form_Load, Save, and validation templates
7. **(3 min)** Test: add a record, save it, check the table

Total: ~20 minutes for a working basic form. Then refine further.

### 12.4 Daily Development Habit

The fastest way to complete all remaining forms and reports:

| Day | Task |
|---|---|
| Day 1 | frm_master dashboard |
| Day 2 | frm_purchases + test |
| Day 3 | frm_debtors + test |
| Day 4 | rpt_sales_daily + rpt_stock_levels |
| Day 5 | rpt_debtor_ageing |
| Day 6 | frm_expenses_obligations |
| Day 7 | frm_returns |
| Day 8 | rpt_profit_loss |
| Day 9 | frm_assets + frm_liabilities |
| Day 10 | frm_drawings + frm_reports_hub |
| Day 11 | rpt_sales_monthly + rpt_cash_flow |
| Day 12 | rpt_balance_sheet + cleanup |

**At 2–3 productive hours per day, your database is complete in 2 weeks.**

---

*Document prepared for Kiyabo Duka — v0.031*  
*Generated May 2026*  
*Revise this document as each item is built and marked ✅*
