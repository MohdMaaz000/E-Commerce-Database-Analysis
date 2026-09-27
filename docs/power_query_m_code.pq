// ============================================================
// Power Query M Code — Sales Data Import & Transformation
// ============================================================
// This query imports all monthly CSV files from the data/ folder,
// combines them, and performs data cleaning & transformation.
//
// HOW TO USE:
// 1. Open Excel → Data → Get Data → From File → From Folder
// 2. Navigate to the "data" folder in this project
// 3. Click "Combine & Transform Data"
// 4. In the Power Query Editor, replace the generated M code
//    with the code below (or use it as a reference)
// 5. Click "Close & Load"
// ============================================================

let
    // Step 1: Import all CSV files from the data folder
    Source = Folder.Files(".\data"),

    // Step 2: Filter to only CSV files (exclude customers.csv and products.csv)
    FilteredFiles = Table.SelectRows(Source, each
        Text.StartsWith([Name], "sales_") and Text.EndsWith([Name], ".csv")
    ),

    // Step 3: Read each CSV file content
    AddContent = Table.AddColumn(FilteredFiles, "ParsedCSV", each
        Csv.Document([Content], [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None])
    ),

    // Step 4: Expand the parsed CSV tables
    ExpandedTable = Table.ExpandTableColumn(AddContent, "ParsedCSV",
        {"Column1","Column2","Column3","Column4","Column5","Column6","Column7",
         "Column8","Column9","Column10","Column11","Column12","Column13",
         "Column14","Column15"}
    ),

    // Step 5: Promote first row as headers
    PromotedHeaders = Table.PromoteHeaders(ExpandedTable, [PromoteAllScalars=true]),

    // Step 6: Remove unnecessary file metadata columns
    RemovedColumns = Table.RemoveColumns(PromotedHeaders,
        {"Content", "Name", "Extension", "Date accessed",
         "Date modified", "Date created", "Attributes", "Folder Path"},
        MissingField.Ignore
    ),

    // Step 7: Clean blank/null values
    ReplacedNulls_Discount = Table.ReplaceValue(RemovedColumns, null, "0",
        Replacer.ReplaceValue, {"Discount"}),
    ReplacedNulls_Payment = Table.ReplaceValue(ReplacedNulls_Discount, null, "Unknown",
        Replacer.ReplaceValue, {"Payment_Method"}),
    ReplacedBlanks_Customer = Table.ReplaceValue(ReplacedNulls_Payment, "", "Unknown",
        Replacer.ReplaceText, {"Customer_ID"}),

    // Step 8: Remove duplicate Order_IDs
    RemovedDuplicates = Table.Distinct(ReplacedBlanks_Customer, {"Order_ID"}),

    // Step 9: Standardize Region names (trim whitespace + Title Case)
    CleanedRegion = Table.TransformColumns(RemovedDuplicates,
        {{"Region", each Text.Trim(Text.Proper(_)), type text}}
    ),

    // Step 10: Convert dates to proper Date type
    ConvertedDates = Table.TransformColumns(CleanedRegion,
        {{"Order_Date", each
            try Date.From(DateTime.From(_))
            otherwise try Date.FromText(_, [Format="yyyy-MM-dd"])
            otherwise try Date.FromText(_, [Format="dd-MM-yyyy"])
            otherwise try Date.FromText(_, [Format="MM/dd/yyyy"])
            otherwise null,
            type date
        }}
    ),

    // Step 11: Remove rows with null dates
    FilteredValidDates = Table.SelectRows(ConvertedDates,
        each [Order_Date] <> null
    ),

    // Step 12: Convert numeric columns to proper types
    ConvertedTypes = Table.TransformColumnTypes(FilteredValidDates, {
        {"Quantity", Int64.Type},
        {"Unit_Price", type number},
        {"Discount", type number},
        {"Sales_Amount", type number},
        {"Cost", type number},
        {"Profit", type number}
    }),

    // Step 13: Add calculated columns
    AddedMonth = Table.AddColumn(ConvertedTypes, "Month",
        each Date.MonthName([Order_Date]), type text),
    AddedMonthNum = Table.AddColumn(AddedMonth, "Month_Num",
        each Date.Month([Order_Date]), Int64.Type),
    AddedQuarter = Table.AddColumn(AddedMonthNum, "Quarter",
        each "Q" & Text.From(Date.QuarterOfYear([Order_Date])), type text),

    // Step 14: Calculate Revenue (recalculate to ensure consistency)
    AddedRevenue = Table.AddColumn(AddedQuarter, "Calculated_Revenue",
        each [Quantity] * [Unit_Price] * (1 - [Discount]), type number),

    // Step 15: Calculate Profit Margin
    AddedProfitMargin = Table.AddColumn(AddedRevenue, "Profit_Margin_%",
        each if [Sales_Amount] > 0
             then ([Profit] / [Sales_Amount]) * 100
             else 0,
        type number),

    // Final: Reorder columns for clean output
    ReorderedColumns = Table.ReorderColumns(AddedProfitMargin, {
        "Order_ID", "Order_Date", "Month", "Month_Num", "Quarter",
        "Customer_ID", "Product_ID", "Region", "State", "Sales_Channel",
        "Quantity", "Unit_Price", "Discount", "Sales_Amount",
        "Calculated_Revenue", "Cost", "Profit", "Profit_Margin_%",
        "Payment_Method", "Order_Status"
    })
in
    ReorderedColumns


// ============================================================
// ADDITIONAL QUERY: Product Lookup Merge
// ============================================================
// This query can be used to merge product details into the
// sales data instead of using XLOOKUP formulas.

/*
let
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    ProductTable = Csv.Document(File.Contents(".\data\products.csv"),
        [Delimiter=",", Encoding=65001]),
    PromotedProducts = Table.PromoteHeaders(ProductTable),
    MergedProducts = Table.NestedJoin(Source, {"Product_ID"},
        PromotedProducts, {"Product_ID"}, "ProductLookup",
        JoinKind.LeftOuter),
    ExpandedProducts = Table.ExpandTableColumn(MergedProducts,
        "ProductLookup",
        {"Product_Name", "Category", "Subcategory", "Brand", "Standard_Cost"})
in
    ExpandedProducts
*/


// ============================================================
// ADDITIONAL QUERY: Customer Lookup Merge
// ============================================================
/*
let
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    CustomerTable = Csv.Document(File.Contents(".\data\customers.csv"),
        [Delimiter=",", Encoding=65001]),
    PromotedCustomers = Table.PromoteHeaders(CustomerTable),
    MergedCustomers = Table.NestedJoin(Source, {"Customer_ID"},
        PromotedCustomers, {"Customer_ID"}, "CustomerLookup",
        JoinKind.LeftOuter),
    ExpandedCustomers = Table.ExpandTableColumn(MergedCustomers,
        "CustomerLookup",
        {"Customer_Name", "Customer_Segment", "City"})
in
    ExpandedCustomers
*/
