# Inventory Accounting System - Design Documentation

## Overview

This system is a comprehensive inventory management and accounting solution built in MS Access with VBA modules. It tracks all inventory movements, calculates stock quantities, and performs GAAP-compliant accounting calculations including COGS, gross profit, and inventory valuation.

## Database Structure

Based on code analysis, the database contains the following tables/queries:

### Transaction Tables/Queries:
- **qry_purchases** - Purchase transactions (fields: `product_spec_id`, `quantity`, `amount`, `purchase_date`, `month_id`)
- **tbl_sales** - Cash sales transactions (fields: `product_spec_id`, `quantity`, `amount`, `sale_date`, `month_id`)
- **qry_sales** - Cash sales query (aggregated view)
- **tbl_debts** - Credit sales / Accounts Receivable (fields: `product_spec_id`, `quantity`, `amount`, `sale_date`, `month_id`)
- **qry_return_inwards** - Sales returns (customers returning goods)
- **qry_return_outwards** - Purchase returns (returning goods to suppliers)
- **qry_sales_office_use** - Internal consumption transactions
- **qry_drawings** - Owner withdrawals
- **tbl_expenses** - Operating expenses (fields: `expense_type_id`, `amount`, `expense_date`)
- **qry_debts_outstanding** - Outstanding receivables calculation

### Key Fields:
- `product_spec_id` - Primary product identifier
- `quantity` - Physical quantity moved
- `amount` - Monetary value of transaction
- `sale_date` / `purchase_date` - Transaction dates
- `month_id` - Period identifier

## Module Architecture

### 1. modStats.vba - Data Access Layer

**Purpose:** Basic aggregation functions for quantities and monetary amounts

**Key Functions:**

#### Monetary Functions (Return Currency):
- `getSales(product_spec_id, product_id, brand_id, category_id, type_id, month_id, start_date, end_date, returnRecordset)` - Cash sales revenue from `qry_sales`
- `getReturnInwards(...)` - Sales returns revenue from `qry_return_inwards`
- `getNetSales(...)` - **Cash Sales - Return Inwards** (Note: Does NOT include Credit Sales)
- `getPurchases(...)` - Purchase costs from `qry_purchases`
- `getReturnOutwards(...)` - Purchase returns from `qry_return_outwards`
- `getNetPurchases(...)` - Purchases - Return Outwards
- `getExpenses(expense_type_id, month_id, start_date, end_date)` - Operating expenses from `tbl_expenses`
- `getReceivablesOutstanding(start_date, end_date)` - Outstanding credit sales from `qry_debts_outstanding`

#### Quantity Functions (Return Long):
- `getPurchasedQuantity(product_spec_id, product_id, up_to_date)` - Uses `qry_purchases`, filters by `purchase_date`
- `getSoldQuantity(...)` - Uses `qry_sales`, filters by `sale_date`
- `getOfficeUseQuantity(...)` - Uses `qry_sales_office_use`, filters by `sale_date`
- `getDrawingsQuantity(...)` - Uses `qry_drawings`, filters by `sale_date`
- `getReturnInwardsQuantity(...)` - Uses `qry_return_inwards`, filters by `sale_date`
- `getReturnOutwardsQuantity(...)` - Uses `qry_return_outwards`, filters by `sale_date` (Note: uses sale_date field)
- `getCreditSalesQuantity(...)` - Uses `tbl_debts`, filters by `sale_date`
- `getCurrentStockQuantity(product_spec_id, product_id, up_to_date)` - Net stock quantity calculation using all transaction types

**Implementation Details:**
- **Monetary functions:** Use DAO.Recordset with SQL queries, iterate through records to sum amounts
- **Quantity functions:** Use `DSum()` function with condition strings
- **Optional `returnRecordset` parameter:** Available in `getSales()`, `getReturnInwards()`, `getPurchases()`, `getReturnOutwards()` - returns recordset instead of total
- **Filtering support:** product_spec_id, product_id, brand_id, category_id, type_id, month_id, date ranges (start_date/end_date), or up_to_date
- **Product ID precedence:** When both product_spec_id and product_id are provided, product_id takes precedence (aggregates across all product_spec_ids)
- **Date handling:** Uses Access date format `#mm/dd/yyyy#` in SQL queries
- **Null handling:** Uses `Nz()` function to return 0 for null values

**Characteristics:**
- Simple SQL/DSum aggregations - no business logic
- Flexible filtering by product, brand, category, type, month, or date range
- No cost calculations or valuation logic
- Returns raw totals from database
- **Important:** `getNetSales()` in modStats only includes Cash Sales - Return Inwards (does NOT include Credit Sales)

### 2. mosAccountings.vba - Business Logic Layer

**Purpose:** GAAP-compliant inventory accounting and valuation using Weighted Average Costing

**Key Functions:**

#### Item-Level Quantity Functions (Return Double):
All use `DSum()` with condition strings. Support: product_spec_id, product_id, month_id, startDate/endDate, or upToDate.
- `getItemQtyPurchased(product_spec_id, product_id, month_id, startDate, endDate, upToDate)` - Uses `qry_purchases`, filters by `purchase_date`
- `getItemQtySold(...)` - Uses `tbl_sales` (Note: different from modStats which uses `qry_sales`), filters by `sale_date`
- `getItemQtyCreditSales(...)` - Uses `tbl_debts`, filters by `sale_date`
- `getItemQtyReturnedInwards(...)` - Uses `qry_return_inwards`, filters by `sale_date`
- `getItemQtyReturnedOutwards(...)` - Uses `qry_return_outwards`, filters by `sale_date` (Note: uses sale_date field)
- `getItemQtyOfficeUse(...)` - Uses `qry_sales_office_use`, filters by `sale_date`
- `getItemQtyDrawings(...)` - Uses `qry_drawings`, filters by `sale_date`
- `getItemInventoryQty(...)` - Net inventory quantity = Purchases + ReturnInwards - Sales - CreditSales - ReturnOutwards - OfficeUse - Drawings

#### Item-Level Value Functions (Return Currency):
All use `DSum()` with condition strings. Support same filtering options as quantity functions.
- `getItemPurchaseValue(...)` - Total purchase costs from `qry_purchases`
- `getItemSalesRevenue(...)` - Cash sales revenue from `qry_sales`
- `getItemCreditSalesRevenue(...)` - Credit sales revenue from `tbl_debts`
- `getItemReturnInwardsValue(...)` - Sales return values from `qry_return_inwards`
- `getItemReturnOutwardsValue(...)` - Purchase return values from `qry_return_outwards`

#### Inventory Valuation Functions:
- `getItemWeightedAvgCost(product_spec_id, upToDate, product_id)` - **Core costing method**
  - Formula: `Total Purchase Cost ÷ Total Purchase Quantity` (up to date)
  - Calculates cumulative weighted average from all purchases up to specified date
  - Returns 0 if total quantity is 0
- `getItemOpeningQty(product_spec_id, periodStartDate, product_id)` - Opening stock quantity
  - **Date Logic:** Calculates stock as of `periodStartDate - 1` (end of previous day = start of period)
- `getItemClosingQty(product_spec_id, periodEndDate, product_id)` - Closing stock quantity
  - Calculates stock as of periodEndDate
- `getItemOpeningStock(product_spec_id, product_id, month_id, startDate, endDate)` - Opening stock VALUE
  - Determines periodStartDate from startDate or month_id
  - Gets opening quantity using `getItemOpeningQty(periodStartDate - 1)`
  - Values at weighted average cost as of `periodStartDate - 1`
  - Returns 0 if opening quantity <= 0
- `getItemClosingStock(product_spec_id, product_id, month_id, startDate, endDate)` - Closing stock VALUE
  - Determines periodEndDate from endDate (with special handling for trailing # or quotes), month_id, or current date
  - Gets closing quantity using `getItemClosingQty(periodEndDate)`
  - Values at weighted average cost as of periodEndDate
  - Returns 0 if closing quantity <= 0

#### Accounting Functions (Return Currency):
- `getItemNetPurchases(...)` - Purchases - Purchase Returns (for period)
- `getItemNetSales(...)` - **Cash Sales + Credit Sales - Sales Returns** (Note: Includes Credit Sales, unlike modStats)
- `getItemCOGS(...)` - Cost of Goods Sold
  - **Calculation Method:** `Opening Stock + Net Purchases - Closing Stock` (Inventory Equation Method)
  - This is the correct GAAP method for weighted average costing
  - Automatically includes all inventory outflows (Sales, Credit Sales, Office Use, Drawings) at weighted average cost
- `getItemGrossProfit(...)` - Net Sales - COGS
- `getItemGrossMargin(...)` - Gross Profit Margin % = (Gross Profit / Net Sales) × 100, rounded to 2 decimals
- `getItemCOGAS(...)` - Cost of Goods Available for Sale = Opening Stock + Net Purchases

#### Aggregate Functions (All Products - Return Currency):
These functions iterate over all product_spec_ids using `getAllProductSpecIDs()` private function:
- `getAllProductSpecIDs()` - Private function that queries `qry_purchases` for distinct product_spec_ids
- `getOpeningStock(month_id, startDate, endDate)` - Sums `getItemOpeningStock()` for all products
- `getClosingStock(...)` - Sums `getItemClosingStock()` for all products
- `getPurchases(...)` - Sums `getItemPurchaseValue()` for all products
- `getNetPurchases(...)` - Sums `getItemNetPurchases()` for all products
- `getNetSales(...)` - Sums `getItemNetSales()` for all products (includes Credit Sales)
- `getCOGS(...)` - Sums `getItemCOGS()` for all products
- `getGrossProfit(...)` - Calculated as `getNetSales() - getCOGS()` (not sum of individual gross profits)
- `getGrossMargin(...)` - Calculated from aggregate `getGrossProfit()` and `getNetSales()`
- `getCOGAS(...)` - Sums `getItemCOGAS()` for all products

#### Validation Function:
- `ValidateItemInventoryEquation(product_spec_id, product_id, month_id, startDate, endDate)` - Returns Boolean
  - Validates: `Opening Stock + Net Purchases = COGS + Closing Stock`
  - Allows 0.01 rounding tolerance
  - Provides detailed debug output

**Implementation Details:**
- **All functions use DSum():** More efficient than recordset iteration for aggregations
- **Date filtering priority:** upToDate > (startDate AND endDate) > month_id
- **Product ID precedence:** When both product_spec_id and product_id provided, product_id takes precedence
- **Opening stock date logic:** Uses `periodStartDate - 1` because opening stock = closing stock of previous day
- **Closing stock date parsing:** Special handling to remove trailing `#` or quotes from endDate parameter
- **COGS calculation:** Uses inventory equation method (not direct quantity × cost calculation) for weighted average costing
- **Debug output:** Extensive Debug.Print statements for troubleshooting

**Characteristics:**
- Uses weighted average costing method (GAAP-compliant)
- Period-based calculations (opening/closing stock)
- Validates inventory equation: Opening + Net Purchases = COGS + Closing
- Professional accounting principles with proper date handling
- **Important:** `getItemNetSales()` includes Credit Sales (unlike modStats.getNetSales())

## Stock Calculation Logic

### Formula:
```
Current Stock Quantity = 
    Purchases 
  + Return Inwards (sales returns - inventory increases)
  - Sales (cash)
  - Credit Sales (debts/accounts receivable)
  - Return Outwards (purchase returns - inventory decreases)
  - Office Use (internal consumption)
  - Drawings (owner withdrawals)
```

### Transaction Type Impact:

| Transaction | Inventory Effect | Revenue/Cost Impact | Included in COGS? |
|------------|------------------|---------------------|-------------------|
| **Purchases** | + Increases | Increases purchase cost | No (inventory asset) |
| **Sales (Cash)** | - Decreases | Increases sales revenue | Yes |
| **Credit Sales (Debts)** | - Decreases | Increases sales revenue + receivables | Yes |
| **Return Inwards** | + Increases | Reduces sales revenue | No (reverses sale) |
| **Return Outwards** | - Decreases | Reduces purchase cost | No (reverses purchase) |
| **Office Use** | - Decreases | Expensed as operating cost | Yes |
| **Drawings** | - Decreases | Owner's equity reduction | Yes |

## Accounting Integration

### Income Statement Components:

#### 1. Net Purchases
```
Net Purchases = Purchases - Return Outwards (Purchase Returns)
```
- Reduces the total cost of goods acquired
- Implemented: `getItemNetPurchases()` / `getNetPurchases()`

#### 2. Net Sales
```
Net Sales = Cash Sales + Credit Sales - Return Inwards (Sales Returns)
```
- Total revenue after accounting for returns
- Credit sales are included in revenue (even though payment is deferred)
- Implemented: `getItemNetSales()` / `getNetSales()`

#### 3. Cost of Goods Sold (COGS)
```
COGS = Opening Stock + Net Purchases - Closing Stock
```
- **Calculation Method:** Uses the Inventory Equation Method (not direct quantity × cost)
- This is the correct GAAP method for weighted average costing because:
  - Weighted average cost changes with each purchase
  - Opening and closing stock are valued at different weighted average costs
  - The equation automatically accounts for all inventory outflows (Sales, Credit Sales, Office Use, Drawings)
- **Important:** Office Use and Drawings ARE included in COGS because they represent inventory that has left the business
- Uses weighted average costing method for historical cost basis
- Implemented: `getItemCOGS()` / `getCOGS()`

#### 4. Gross Profit
```
Gross Profit = Net Sales - COGS
```
- Basic profitability metric
- Implemented: `getItemGrossProfit()` / `getGrossProfit()`

#### 5. Gross Margin
```
Gross Margin % = (Gross Profit / Net Sales) × 100
```
- Profitability percentage
- Implemented: `getItemGrossMargin()` / `getGrossMargin()`

### Balance Sheet Components:

#### Opening Stock Value
```
Opening Stock = Opening Quantity × Weighted Avg Cost (as of period start)
```
- **Date Logic:** Opening stock for a period = Closing stock from previous period
  - Example: If period starts on 01/07/2025, opening stock = stock at end of 01/06/2025
  - Implementation: `upToDate = periodStartDate - 1` (end of previous day = start of period)
- Calculated at the beginning of a reporting period
- Uses weighted average cost from the day before period start (`periodStartDate - 1`)
- Returns 0 if opening quantity <= 0
- Implemented: `getItemOpeningStock()` / `getOpeningStock()`

#### Closing Stock Value
```
Closing Stock = Closing Quantity × Weighted Avg Cost (as of period end)
```
- Calculated at the end of a reporting period
- **Date Determination:** Uses endDate (with special parsing to remove trailing # or quotes), month_id (last day of month), or current date
- Uses weighted average cost as of period end date
- Returns 0 if closing quantity <= 0
- Implemented: `getItemClosingStock()` / `getClosingStock()`

### Key Accounting Equation:

```
Opening Stock + Net Purchases = COGS + Closing Stock
```

This equation must always balance:
- **Left Side (Availability):** What we had + what we bought
- **Right Side (Disposition):** What we sold/used + what we have left

Validated by: `ValidateItemInventoryEquation()`

## How Each Transaction Type is Integrated

### 1. Drawings (Owner Withdrawals)

**Stock Impact:** Decreases inventory quantity

**Integration:**
- Included in `getCurrentStockQuantity()` calculation (subtracted)
- Included in COGS calculation (multiplied by weighted average cost)
- Reduces owner's equity (drawings account)
- Not included in sales revenue (it's not a sale)
- **Formula Position:**
  - Stock: `- Drawings`
  - COGS: `+ (Drawings Qty × Avg Cost)`

### 2. Returns Inwards (Sales Returns)

**Stock Impact:** Increases inventory quantity (goods come back)

**Integration:**
- Included in `getCurrentStockQuantity()` calculation (added)
- Reduces Net Sales: `Net Sales = Sales - Returns Inwards`
- Not included in COGS (reverses the original sale)
- **Formula Position:**
  - Stock: `+ Return Inwards`
  - Net Sales: `- Return Inwards Value`
  - COGS: Not included (reverses original COGS)

### 3. Returns Outwards (Purchase Returns)

**Stock Impact:** Decreases inventory quantity (goods sent back to supplier)

**Integration:**
- Included in `getCurrentStockQuantity()` calculation (subtracted)
- Reduces Net Purchases: `Net Purchases = Purchases - Returns Outwards`
- Reduces purchase cost basis
- **Formula Position:**
  - Stock: `- Return Outwards`
  - Net Purchases: `- Return Outwards Value`
  - COGS: Not directly (but affects weighted average cost calculation)

### 4. Credit Sales (Debts / Accounts Receivable)

**Stock Impact:** Decreases inventory quantity (same as cash sales)

**Integration:**
- Included in `getCurrentStockQuantity()` calculation (subtracted)
- Included in Net Sales: `Net Sales = Cash Sales + Credit Sales - Returns`
- Included in COGS: `COGS = (Sales + Credit Sales + ...) × Avg Cost`
- Creates accounts receivable (asset)
- **Formula Position:**
  - Stock: `- Credit Sales`
  - Net Sales: `+ Credit Sales Value`
  - COGS: `+ (Credit Sales Qty × Avg Cost)`

### 5. Office Use (Internal Consumption)

**Stock Impact:** Decreases inventory quantity

**Integration:**
- Included in `getCurrentStockQuantity()` calculation (subtracted)
- Included in COGS calculation (goods used internally)
- Expensed as operating cost
- **Formula Position:**
  - Stock: `- Office Use`
  - COGS: `+ (Office Use Qty × Avg Cost)`
  - Not included in sales revenue

## Weighted Average Costing Method

The system uses **Weighted Average Costing** (also called Average Cost Method):

```
Weighted Average Cost = Total Purchase Cost ÷ Total Purchase Quantity (up to date)
```

**Characteristics:**
- Calculated cumulatively from all purchases up to a given date
- Used for all inventory valuations (opening stock, closing stock, COGS)
- Provides historical cost basis (GAAP-compliant)
- Implemented: `getItemWeightedAvgCost()`

**Example:**
- Purchase 1: 100 units @ $10 = $1,000
- Purchase 2: 50 units @ $12 = $600
- Weighted Avg Cost = ($1,000 + $600) ÷ (100 + 50) = $10.67/unit

## Example Calculation

Based on the test data from SOME.TXT for product_spec_id 26:

**Quantities:**
- Purchased: 150
- Sold (cash): 41
- Office Use: 7
- Returned Inwards: 2
- Returned Outwards: 5
- Credit Sales: 8
- Drawings: 3

**Stock Calculation:**
```
Current Stock = 150 + 2 - 41 - 8 - 5 - 7 - 3 = 88 units
```

**Accounting Integration:**
1. **Net Purchases** = Purchase Value - Return Outwards Value
2. **Net Sales** = (Cash Sales + Credit Sales) - Return Inwards Value
3. **COGS** = (41 + 8 + 7 + 3) × Weighted Avg Cost = 59 × Avg Cost
4. **Gross Profit** = Net Sales - COGS

## Design Patterns

1. **Separation of Concerns**
   - **modStats:** Data retrieval and basic aggregation (no business logic)
   - **mosAccountings:** Business logic and accounting calculations (GAAP-compliant)

2. **Consistent Naming Convention**
   - `getItem*()` - Single product functions (requires `product_spec_id` as first parameter)
   - `get*()` - Aggregate functions (all products, iterates over product_spec_ids)
   - Both modules follow this pattern consistently

3. **Flexible Date Filtering**
   - **Priority order:** `upToDate` > (`startDate` AND `endDate`) > `month_id`
   - **modStats:** Supports date ranges (start_date/end_date) or up_to_date for quantity functions
   - **mosAccountings:** Supports all three: upToDate, startDate/endDate range, or month_id
   - Consistent parameter patterns across functions

4. **Database Abstraction**
   - Uses queries (`qry_*`) for complex joins/views
   - Uses tables (`tbl_*`) for direct data access
   - **Note:** modStats uses `qry_sales` for getSoldQuantity(), while mosAccountings uses `tbl_sales` for getItemQtySold()

5. **Null Handling**
   - Uses `Nz()` function to handle null values (returns 0)
   - Prevents calculation errors
   - Applied consistently across all aggregation functions

6. **Product ID Precedence**
   - When both `product_spec_id` and `product_id` are provided, `product_id` takes precedence
   - Aggregates across all product_spec_ids for the given product_id
   - Applied consistently in both modules

7. **Data Access Methods**
   - **modStats monetary functions:** DAO.Recordset with SQL queries (can return recordset or total)
   - **modStats quantity functions:** DSum() with condition strings
   - **mosAccountings all functions:** DSum() with condition strings (more efficient)

## Validation and Error Handling

- `ValidateItemInventoryEquation()` - Validates accounting equation balance
- Rounding tolerance: 0.01 for currency comparisons
- Debug output for validation diagnostics
- Warning messages for missing cost data in COGS calculations

## Usage Notes

- **Return Types:**
  - Monetary functions return `Currency` type
  - Quantity functions return `Double` (mosAccountings) or `Long` (modStats)
  - Boolean for validation functions

- **Date Parameters:**
  - Use Access date format: `#mm/dd/yyyy#` in SQL queries
  - Can pass Date variables directly in VBA
  - Special handling in `getItemClosingStock()` for endDate with trailing # or quotes

- **Function Usage:**
  - Functions are designed to work standalone or in aggregate calculations
  - Period-based calculations require either `month_id` or date range parameters (startDate/endDate)
  - `upToDate` parameter is useful for cumulative calculations up to a specific point in time

- **Important Differences Between Modules:**
  - **modStats.getNetSales():** Cash Sales - Return Inwards only (does NOT include Credit Sales)
  - **mosAccountings.getItemNetSales():** Cash Sales + Credit Sales - Return Inwards (includes Credit Sales)
  - **modStats.getSoldQuantity():** Uses `qry_sales`
  - **mosAccountings.getItemQtySold():** Uses `tbl_sales`
  - **modStats:** Can return recordset for some monetary functions (returnRecordset parameter)
  - **mosAccountings:** All functions return aggregated values only

- **COGS Calculation:**
  - Uses Inventory Equation Method: `Opening Stock + Net Purchases - Closing Stock`
  - This is the correct method for weighted average costing (not direct quantity × cost)
  - Automatically accounts for all inventory outflows at weighted average cost

