# Transaction Integration Explanation

## Direct Answers to Questions from SOME.TXT

### Q: What is stock if we bought, returned some outwards, sold and returned inwards some, has been drawn, has been debted, and office used?

**Answer:**
```
Current Stock Quantity = 
    Purchases 
  + Return Inwards (customers returning goods to you)
  - Sales (cash)
  - Credit Sales (debts - goods sold on credit)
  - Return Outwards (you returning goods to suppliers)
  - Office Use (internal consumption)
  - Drawings (owner withdrawals)
```

**Example using your test data (product_spec_id 26):**
- Purchased: 150
- Sold: 41
- Office Use: 7
- Returned Inwards: 2
- Returned Outwards: 5
- Credit Sales: 8
- Drawings: 3

**Calculation:**
```
Stock = 150 + 2 - 41 - 8 - 5 - 7 - 3 = 88 units
```

### Q: How did you integrate them to sales, purchases, gross/profit or purchase?

**Answer:**

#### 1. **Net Purchases** (Income Statement - Cost of Goods)
```
Net Purchases = Purchases - Return Outwards
```
- Return Outwards **reduces** purchase cost because you're sending goods back to suppliers
- Example: If you purchased $10,000 and returned $500, Net Purchases = $9,500

#### 2. **Net Sales** (Income Statement - Revenue)
```
Net Sales = Cash Sales + Credit Sales - Return Inwards
```
- Credit Sales are **included** because revenue is recognized when goods are sold (even if payment is later)
- Return Inwards **reduces** sales because customers are returning goods
- Example: Cash Sales $5,000 + Credit Sales $2,000 - Returns $300 = Net Sales $6,700

#### 3. **COGS (Cost of Goods Sold)** (Income Statement)
```
COGS = Opening Stock + Net Purchases - Closing Stock
```
- **Calculation Method:** Uses the Inventory Equation Method (not direct quantity × cost)
- This is the correct GAAP method for weighted average costing because:
  - Weighted average cost changes with each purchase
  - Opening and closing stock are valued at different weighted average costs
  - The equation automatically accounts for all inventory outflows
- **All inventory outflows are included** (costed at weighted average purchase price):
  - **Credit Sales** are included (same as cash sales - they reduce inventory)
  - **Office Use** is included (goods consumed internally)
  - **Drawings** are included (owner took goods)
- **Returns are NOT included** because:
  - Return Inwards: Reverses a sale (goods come back, original COGS was reversed)
  - Return Outwards: Reverses a purchase (affects Net Purchases, not COGS)

#### 4. **Gross Profit** (Income Statement)
```
Gross Profit = Net Sales - COGS
```
- Basic profitability: Revenue minus the cost of goods that left inventory

### Q: How did you do these drawings, returns inwards, debts, etc?

**Answer:**

#### **Drawings (Owner Withdrawals)**
- **Stock Impact:** Decreases inventory (subtracted in stock calculation)
- **Accounting Impact:** 
  - Included in COGS: `COGS = (Sales + Credit Sales + Office Use + Drawings) × Avg Cost`
  - Reduces owner's equity (not included in sales revenue)
- **Why in COGS?** Because goods left the business (even though not sold)

#### **Returns Inwards (Sales Returns)**
- **Stock Impact:** Increases inventory (added in stock calculation - goods come back)
- **Accounting Impact:**
  - Reduces Net Sales: `Net Sales = Sales - Returns Inwards`
  - NOT included in COGS (reverses the original sale's COGS)
- **Why reduce sales?** Because the original sale is being reversed

#### **Returns Outwards (Purchase Returns)**
- **Stock Impact:** Decreases inventory (subtracted - goods sent back to supplier)
- **Accounting Impact:**
  - Reduces Net Purchases: `Net Purchases = Purchases - Returns Outwards`
  - NOT included in COGS (reverses the original purchase)
- **Why reduce purchases?** Because the original purchase is being reversed

#### **Debts (Credit Sales / Accounts Receivable)**
- **Stock Impact:** Decreases inventory (subtracted - same as cash sales)
- **Accounting Impact:**
  - Included in Net Sales: `Net Sales = Cash Sales + Credit Sales - Returns`
  - Included in COGS: `COGS = (Sales + Credit Sales + ...) × Avg Cost`
  - Creates accounts receivable (asset on balance sheet)
- **Why included in sales?** Revenue is recognized when goods are sold (accrual accounting), not when cash is received

#### **Office Use (Internal Consumption)**
- **Stock Impact:** Decreases inventory (subtracted)
- **Accounting Impact:**
  - Included in COGS: `COGS = (Sales + ... + Office Use + ...) × Avg Cost`
  - Expensed as operating cost
- **Why in COGS?** Because goods left inventory (used internally)

## Summary Table

| Transaction | Stock Formula | Net Sales | Net Purchases | COGS | Balance Sheet |
|------------|--------------|-----------|---------------|------|---------------|
| Purchases | + Add | No | + Add | No | Inventory Asset |
| Sales (Cash) | - Subtract | + Add | No | + Add (qty × cost) | No |
| Credit Sales | - Subtract | + Add | No | + Add (qty × cost) | + Receivables |
| Return Inwards | + Add | - Subtract | No | No (reverses sale) | No |
| Return Outwards | - Subtract | No | - Subtract | No (reverses purchase) | No |
| Office Use | - Subtract | No | No | + Add (qty × cost) | No |
| Drawings | - Subtract | No | No | + Add (qty × cost) | - Equity |

## Complete Accounting Flow

```
INCOME STATEMENT:
─────────────────
Net Sales (Cash + Credit - Returns Inwards)
  - COGS (Sold + Credit + Office Use + Drawings) × Avg Cost
─────────────────
= Gross Profit
```

```
BALANCE SHEET (Asset Side):
──────────────────────────
Inventory:
  Opening Stock Value (Opening Qty × Avg Cost)
  + Net Purchases (Purchases - Returns Outwards)
  - COGS
  = Closing Stock Value (Closing Qty × Avg Cost)

Accounts Receivable:
  Credit Sales - Payments Received
```

## Validation

The system validates that:
```
Opening Stock + Net Purchases = COGS + Closing Stock
```

This ensures all inventory movements are properly accounted for.

**Validation Function:**
- `ValidateItemInventoryEquation()` in mosAccountings.vba
- Returns Boolean (True if difference < 0.01)
- Provides detailed debug output showing all components
- Since COGS is calculated as `Opening + Net Purchases - Closing`, this validation will always pass mathematically, but it's kept for verification and debugging

## Technical Implementation Notes

### Module Differences

**modStats.vba (Data Access Layer):**
- `getNetSales()` = Cash Sales - Return Inwards only (does NOT include Credit Sales)
- Uses `qry_sales` for sales data
- Can return recordset or total for monetary functions
- Simple aggregations with no business logic

**mosAccountings.vba (Business Logic Layer):**
- `getItemNetSales()` = Cash Sales + Credit Sales - Return Inwards (includes Credit Sales)
- Uses `tbl_sales` for sales quantity data
- Always returns aggregated values
- Implements GAAP-compliant accounting calculations

### COGS Calculation Method

The system uses the **Inventory Equation Method** for COGS calculation:
```
COGS = Opening Stock + Net Purchases - Closing Stock
```

This is the correct method for weighted average costing because:
1. Weighted average cost changes with each purchase
2. Opening stock is valued at weighted average cost as of period start
3. Closing stock is valued at weighted average cost as of period end
4. The equation automatically accounts for all inventory outflows (Sales, Credit Sales, Office Use, Drawings) at their respective weighted average costs

This method is more accurate than calculating `(Quantity × Average Cost)` directly because it properly handles the changing cost basis throughout the period.

