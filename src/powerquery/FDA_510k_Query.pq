let
    // --- Configuration ---
    // Get current date/time for dynamic calculation
    Today = Date.From(DateTime.FixedLocalNow()),
    // Set the target month to the *previous* full month relative to today
    PrevMonthDate = Date.AddMonths(Today, -1),
    // Calculate the start and end dates of that target month
    StartDate = Date.StartOfMonth(PrevMonthDate),
    EndDate = Date.EndOfMonth(PrevMonthDate),
    // Format dates into YYYYMMDD text format required by the API URL
    StartDateText = Date.ToText(StartDate, "yyyyMMdd"),
    EndDateText = Date.ToText(EndDate, "yyyyMMdd"),

    // --- Build Dynamic API URL ---
    // Define the base URL for the FDA 510(k) JSON endpoint
    BaseUrl = "https://api.fda.gov/device/510k.json?search=",
    // Construct the date range search parameter using the formatted dates
    DateSearch = "decision_date:[" & StartDateText & "+TO+" & EndDateText & "]",
    // Define the API limit parameter (FDA API max is 1000 per request without pagination)
    // Note: This query assumes fewer than 1000 results per month. Pagination would require more complex M code.
    LimitParam = "&limit=1000",
    // Combine the base URL, search parameters, and limit into the final API URL
    DynamicApiUrl = BaseUrl & DateSearch & LimitParam,

    // --- Fetch and Process Data ---
    // Attempt to fetch data from the constructed API URL
    // Use 'try...otherwise null' to gracefully handle potential errors (network issues, API downtime, invalid response)
    // Set a longer timeout (e.g., 300 seconds = 5 minutes = #duration(0, 0, 5, 0))
    Source = try Json.Document(Web.Contents(DynamicApiUrl, [Timeout=#duration(0, 0, 5, 0)])) otherwise null,

    // Check if the API call was successful (Source is not null) and if the response contains the expected 'results' field
    // If not, provide an empty list to prevent downstream errors
    ResultsList = if Source <> null and Record.HasFields(Source, "results") then Source[results] else {},

    // Convert the list of JSON records into a Power Query table
    // Each item in ResultsList becomes a row with one column named "RawDataRecord" containing the original record
    #"Converted to Table" = Table.FromList(ResultsList, Splitter.SplitByNothing(), {"RawDataRecord"}, null, ExtraValues.Error),

    // --- Expand Necessary Columns & Rename Cleanly ---
    // Define the specific raw field names we want to extract from the JSON records
    RequiredRawColumns = {"k_number", "device_name", "applicant", "decision_date", "date_received",
                          "product_code", "advisory_committee", "submission_type", "city", "state",
                          "country_code", "statement_or_summary", "contact"},
    // Define the desired, user-friendly output column names for Excel/VBA
    OutputColumnNames = {"K_Number", "DeviceName", "Applicant", "DecisionDate", "DateReceived",
                         "PC", "AC", "SubmType", "City", "State", "Country", "Statement", "Contact"},

    // Expand the 'RawDataRecord' column, selecting only the required fields and renaming them simultaneously
    // This makes the query robust against unexpected extra fields in the API response
    #"Expanded Data" = Table.ExpandRecordColumn(#"Converted to Table", "RawDataRecord", RequiredRawColumns, OutputColumnNames),

    // --- Basic Data Type Conversion ---
    // Explicitly set the data types for key columns to ensure correct handling in Power Query and Excel
    #"Changed Types" = Table.TransformColumnTypes(#"Expanded Data",{
        {"DecisionDate", type date}, {"DateReceived", type date}, // Set date columns to Date type
        {"K_Number", type text}, {"DeviceName", type text}, {"Applicant", type text}, // Ensure key identifiers are Text
        {"PC", type text}, {"AC", type text}, {"SubmType", type text}, // Codes as Text
        {"City", type text}, {"State", type text}, {"Country", type text}, // Location info as Text
        {"Statement", type text}, {"Contact", type text} // Other text fields
        }),

    // --- Calculate Processing Time (in days) ---
    // Calculate the difference between DecisionDate and DateReceived in days
    // Use 'try...otherwise null' to handle cases where one or both dates might be missing or invalid
    // Set the result type to Int64.Type (whole number)
    #"Added ProcTime" = Table.AddColumn(#"Changed Types", "ProcTimeDays",
        each try Duration.Days([DecisionDate] - [DateReceived]) otherwise null, Int64.Type),

    // --- Add Direct FDA Link ---
    // Construct a hyperlink to the specific 510(k) record on the FDA's online database
    // Check if K_Number is present before creating the link
    #"Added URL Column" = Table.AddColumn(#"Added ProcTime", "FDA_Link",
        each if [K_Number] <> null and [K_Number] <> "" then
                 "https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfpmn/pmn.cfm?ID=" & [K_Number]
             else
                 null, type text), // Set the result type to Text

    // --- Select and Reorder Columns for Output ---
    // Select only the columns created by this query that are needed as input for the VBA processing
    // Note: The final, precise column order seen in Excel is controlled by the VBA 'ReorganizeColumns' function.
    // This step just provides a logical initial set of columns.
    #"Selected Columns" = Table.SelectColumns(#"Added URL Column", {
        "K_Number",         // Key Identifier
        "DecisionDate",     // Dates
        "DateReceived",
        "ProcTimeDays",
        "Applicant",        // Company Info
        "Contact",
        "DeviceName",       // Device Info
        "Statement",
        "AC",               // Codes & Types
        "PC",
        "SubmType",
        "City",             // Location
        "State",
        "Country",
        "FDA_Link"          // Added Link
    }),

    // --- Sort results by decision date (most recent first) ---
    // Provides a default sort order which might be overridden by VBA later (e.g., sorting by score)
    #"Sorted Rows" = Table.Sort(#"Selected Columns",{{"DecisionDate", Order.Descending}})

// Final step: Output the resulting table
in
    #"Sorted Rows"
