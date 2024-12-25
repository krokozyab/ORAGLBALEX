namespace OraGlBalEx

module OraExcelDna =
    open System
    open ExcelDna.Integration
    open FsHttp
    open System.Text.Json.Serialization
    open Microsoft.FSharp.Control
    open Newtonsoft.Json
    open System.Net

    open Security
    open JsonConverters
    open Helpers

    /// Request output fields
    [<CLIMutable>]
    type BalancesFields =
        { [<JsonPropertyName("AccountGroupName")>]
          AccountGroupName: string option
          [<JsonPropertyName("AccountName")>]
          AccountName: string option
          [<JsonPropertyName("LedgerSetName")>]
          LedgerSetName: string option
          [<JsonPropertyName("LedgerName")>]
          LedgerName: string option
          [<JsonPropertyName("Currency")>]
          Currency: string option
          [<JsonPropertyName("CurrentAccountingPeriod")>]
          CurrentAccountingPeriod: string option
          [<JsonPropertyName("PeriodName")>]
          PeriodName: string option
          [<JsonPropertyName("CurrentPeriodBalance")>]
          CurrentPeriodBalance: FloatOrString option
          [<JsonPropertyName("BudgetBalance")>]
          BudgetBalance: FloatOrString option
          [<JsonPropertyName("Scenario")>]
          Scenario: string option
          [<JsonPropertyName("AccountCombination")>]
          AccountCombination: string option
          [<JsonPropertyName("DetailAccountCombination")>]
          DetailAccountCombination: string option
          [<JsonPropertyName("BeginningBalance")>]
          BeginningBalance: FloatOrString option
          [<JsonPropertyName("PeriodActivity")>]
          PeriodActivity: FloatOrString option
          [<JsonPropertyName("EndingBalance")>]
          EndingBalance: FloatOrString option
          [<JsonPropertyName("AmountType")>]
          AmountType: string option
          [<JsonPropertyName("CurrencyType")>]
          CurrencyType: string option
          [<JsonPropertyName("ErrorDetail")>]
          ErrorDetail: string option }

    [<CLIMutable>]
    type LedgersResponse =
        { [<JsonPropertyName("items")>]
          Items: BalancesFields list
          [<JsonPropertyName("hasMore")>]
          HasMore: bool }


    /// Function to deserialize the JSON string into LedgersResponse using custom settings
    let deserializeLedgersResponse (json: string) : Result<LedgersResponse, string> =
        try
            let response = JsonConvert.DeserializeObject<LedgersResponse>(json, jsonSettings)
            //Some response
            if isNull (response :> obj) then
                Error "Deserialization resulted in null."
            else
                Ok response
        with
        | :? JsonException as ex ->
            let errorMsg = $"JSON Deserialization failed: {ex.Message}"
            printfn "%s" errorMsg
            Error errorMsg
        | ex ->
            let errorMsg = $"Deserialization failed: {ex.Message}"
            printfn "Deserialization failed: %s" ex.Message
            Error errorMsg


    /// Adds a key-value pair to the map if the value is `Some`, handling `string option` types
    let addIfSomeString (key: string) (valueOpt: string option) : (string * obj) option =
        match valueOpt with
        | Some str -> Some(key, box str)
        | None -> None

    /// Adds a key-value pair to the map if the value is `Some`, handling `FloatOrString option` types
    let addIfSomeFloatOrString (key: string) (valueOpt: FloatOrString option) : (string * obj) option =
        match valueOpt with
        | Some (FloatOrString.FloatValue num) -> Some(key, box num)
        | Some (FloatOrString.StringValue str) -> Some(key, box str)
        | None -> None

    /// Converts a BalancesFields record to a list of (string * obj), omitting `None` fields and preserving order
    let balancesFieldsToList (balance: BalancesFields) : (string * obj) list =
        [
          // Handle string option fields
          addIfSomeString "AccountGroupName" balance.AccountGroupName
          addIfSomeString "AccountName" balance.AccountName
          addIfSomeString "LedgerSetName" balance.LedgerSetName
          addIfSomeString "LedgerName" balance.LedgerName
          addIfSomeString "Currency" balance.Currency
          addIfSomeString "CurrentAccountingPeriod" balance.CurrentAccountingPeriod
          addIfSomeString "PeriodName" balance.PeriodName
          addIfSomeString "Scenario" balance.Scenario
          addIfSomeString "AccountCombination" balance.AccountCombination
          addIfSomeString "DetailAccountCombination" balance.DetailAccountCombination
          addIfSomeString "AmountType" balance.AmountType
          addIfSomeString "CurrencyType" balance.CurrencyType
          addIfSomeString "ErrorDetail" balance.ErrorDetail

          // Handle FloatOrString option fields
          addIfSomeFloatOrString "CurrentPeriodBalance" balance.CurrentPeriodBalance
          addIfSomeFloatOrString "BudgetBalance" balance.BudgetBalance
          addIfSomeFloatOrString "BeginningBalance" balance.BeginningBalance
          addIfSomeFloatOrString "PeriodActivity" balance.PeriodActivity
          addIfSomeFloatOrString "EndingBalance" balance.EndingBalance ]
        |> List.choose id // Filters out `None` values, keeping only `Some (key, value)`


    /// Base URL of the Oracle Fusion REST API. REST Server URL. Typically, the URL of your Oracle Cloud service. For example
    /// "https://servername.fa.us2.oraclecloud.com:443"
    let baseAPIUrl = GetSecret("baseAPIUrl")

    /// GL Balances API url
    /// https://docs.oracle.com/en/cloud/saas/financials/24d/farfa/op-ledgerbalances-get.html
    let balancesAPIUrl = "/fscmRestApi/resources/11.13.18.05/ledgerBalances"

    /// Function to perform the HTTP GET request and return the response
    let fetchLedgers
        (requestLimit: int)
        (offset: int)
        (encodedCredentials: string)
        (balancesDisplayFields: string)
        (balancesFinder: string)
        =
        http {
            GET $"{baseAPIUrl}{balancesAPIUrl}"

            query [ "offset", $"{offset}"
                    "limit", $"{requestLimit}"
                    "onlyData", "True"
                    "fields", balancesDisplayFields
                    "finder", balancesFinder ]

            header "Authorization" $"Basic {encodedCredentials}"
            Accept "application/json" // Specify that we expect JSON response
        }
        |> Request.sendAsync

    /// Fetch and deserialize the "items" array
    let fetchAndDeserializeLedgersAsync requestLimit encodedCredentials balancesDisplayFields balancesFinder () =
        let rec helper (offset: int) (listAcc: BalancesFields list) =
            async {
                try
                    let! response =
                        fetchLedgers requestLimit offset encodedCredentials balancesDisplayFields balancesFinder

                    match response.statusCode with
                    | HttpStatusCode.OK ->
                        let! content =
                            response.content.ReadAsStringAsync()
                            |> Async.AwaitTask

                        match deserializeLedgersResponse content with
                        | Ok deserializedResponse ->
                            let fetchedItems = deserializedResponse.Items
                            let hasMore = deserializedResponse.HasMore
                            //printfn "Successfully fetched and deserialized %d balances." (List.length fetchedItems)
                            match hasMore with
                            | true ->
                                let! result = helper (offset + requestLimit) (listAcc @ fetchedItems)
                                return result
                            | _ -> return Ok(listAcc @ fetchedItems) // No more data to fetch
                        | Error errMsg ->
                            // Deserialization failed; return the accumulated list so far
                            printfn "Failed to deserialize response."
                            return Error errMsg
                    | HttpStatusCode.Unauthorized ->
                        let err = "Unauthorized: Check your credentials."
                        printfn "%s" err
                        return Error err
                    | HttpStatusCode.Forbidden ->
                        let err = "Forbidden: You don't have access to this resource."
                        printfn "%s" err
                        return Error err
                    | HttpStatusCode.NotFound ->
                        let err = "Not Found: The requested resource does not exist."
                        printfn "%s" err
                        return Error err
                    | status ->
                        let err = $"Request failed with status code {(int status)}"
                        printfn "%s" err
                        return Error err
                with
                | ex ->
                    let err = $"An error occurred: {ex.Message}"
                    printfn "%s" err
                    return Error err
            }
        // Start the recursion with the initial offset and an empty list
        helper 0 []

    /// Function to split DetailAccountCombination into segments
    let splitDetailAccountCombination (s: string) =
        s.Split([| '.'; '-' |], StringSplitOptions.RemoveEmptyEntries)

    /// Helper Function to safely find the maximum value in a list
    let tryFindMax (list: int list) : int option =
        match list with
        | [] -> None
        | head :: tail -> Some(List.fold (fun acc x -> if x > acc then x else acc) head tail)


    /// Excel-DNA Function to return BalancesFields data as a 2D array with headers at the top and segX columns
    [<ExcelFunction(Description = "Returns BalancesFields data as a two-dimensional array with headers at the top, including segmented DetailAccountCombination.")>]
    let WriteBalancesFieldsToExcel balancesFinder balancesDisplayFields : obj =
        async {
            // Define your credentials and requestLimit
            let requestLimit = 500

            let encodedCredentials =
                (GetSecret "oracleuser", GetSecret "oraclepassword")
                |> encodeBasicAuth

            // Fetch and deserialize balances asynchronously
            let! fetchResult =
                fetchAndDeserializeLedgersAsync requestLimit encodedCredentials balancesDisplayFields balancesFinder ()

            match fetchResult with
            | Error errMsg ->
                // Return the error message to Excel
                return errMsg :> obj
            | Ok dataList ->
                if List.isEmpty dataList then
                    // Return a message indicating no data was fetched
                    return "No data fetched." :> obj
                else
                    // Proceed to process and return data as 2D array
                    let processedData = dataList |> List.map balancesFieldsToList

                    // Extract headers from the first record to preserve order
                    let headers =
                        match processedData with
                        | firstRecord :: _ -> firstRecord |> List.map fst
                        | [] -> []

                    // Determine the maximum number of segments across all records
                    // Exclude records where DetailAccountCombination is 'N/A'
                    let maxSegments =
                        processedData
                        |> List.choose (fun record ->
                            match List.tryFind (fun (k, _) -> k = "DetailAccountCombination") record with
                            | Some (_, value) ->
                                match value with
                                | :? string as s when s <> "N/A" ->
                                    let segments = splitDetailAccountCombination s
                                    Some segments.Length
                                | _ -> None
                            | None -> None)
                        |> tryFindMax
                        |> Option.defaultValue 0

                    // Generate 'seg1' to 'segN' headers
                    let segHeaders =
                        [ 1..maxSegments ]
                        |> List.map (fun i -> sprintf "seg%d" i)

                    // Find the index of 'DetailAccountCombination' in headers
                    let detailIndexOpt =
                        headers
                        |> List.tryFindIndex (fun h -> h = "DetailAccountCombination")

                    // Insert 'segX' headers after 'DetailAccountCombination'
                    let updatedHeaders =
                        match detailIndexOpt with
                        | Some idx ->
                            let before = headers[0..idx]

                            let after =
                                if idx + 1 < headers.Length then
                                    headers[idx + 1 ..]
                                else
                                    []

                            before @ segHeaders @ after
                        | None ->
                            // If 'DetailAccountCombination' not found, append 'segX' at the end
                            headers @ segHeaders

                    // Prepare data rows aligned with updated headers
                    let rows =
                        processedData
                        |> List.map (fun record ->
                            // Convert the list to a map for easy access
                            let recordMap = record |> Map.ofList

                            // Extract existing row data
                            let rowData =
                                headers
                                |> List.map (fun header ->
                                    match Map.tryFind header recordMap with
                                    | Some value -> value
                                    | None -> box "")
                                |> Array.ofList

                            // Extract and split 'DetailAccountCombination' segments
                            let segments =
                                match Map.tryFind "DetailAccountCombination" recordMap with
                                | Some (:? string as s) when s <> "N/A" ->
                                    splitDetailAccountCombination s |> Array.map box
                                | _ -> Array.empty // If 'N/A' or not present, leave segments empty

                            // Pad segments with empty strings if necessary
                            let paddedSegments =
                                if segments.Length < maxSegments then
                                    Array.append segments (Array.create (maxSegments - segments.Length) (box ""))
                                else
                                    segments[0 .. maxSegments - 1]

                            // Insert the segments into the row data after 'DetailAccountCombination'
                            match detailIndexOpt with
                            | Some idx ->
                                // Split the rowData into before and after the DetailAccountCombination
                                let before = rowData[0..idx]

                                let after =
                                    if idx + 1 < rowData.Length then
                                        rowData[idx + 1 ..]
                                    else
                                        [||]
                                // Concatenate before, segments, and after
                                Array.concat [ before
                                               paddedSegments
                                               after ]
                            | None ->
                                // If 'DetailAccountCombination' not found, append segments at the end
                                Array.append rowData paddedSegments)
                        |> List.toArray

                    // Convert headers to obj[] by boxing each string
                    let headersObjArray: obj [] = updatedHeaders |> List.toArray |> Array.map box

                    // Combine headers and rows into obj[][]
                    let allData: obj [] [] = Array.append [| headersObjArray |] rows

                    // Convert to a 2D object array
                    let rowCount = Array.length allData
                    let colCount = updatedHeaders.Length
                    let resultArray = Array2D.init rowCount colCount (fun i j -> allData[i][j])

                    // Return the two-dimensional array as an object for Excel-DNA
                    return (box resultArray)
        }
        |> Async.RunSynchronously
