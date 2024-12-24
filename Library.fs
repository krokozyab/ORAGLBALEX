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

    /// Global error message
    let mutable errMSG = "Unknown error"


    /// Function to deserialize the JSON string into LedgersResponse using custom settings
    let deserializeLedgersResponse (json: string) : LedgersResponse option =
        try
            let response = JsonConvert.DeserializeObject<LedgersResponse>(json, jsonSettings)
            Some response
        with
        | :? JsonException as ex ->
            printfn "JSON Deserialization failed: %s" ex.Message
            errMSG <- "JSON Deserialization failed: " + ex.Message
            None
        | ex ->
            printfn "Deserialization failed: %s" ex.Message
            errMSG <- "Deserialization failed: " + ex.Message
            None


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

    /// Base URL of the Oracle Fusion REST API. REST Server URL. Typically, the URL of your Oracle Cloud service. For example
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

    /// Asynchronous function to fetch and deserialize the "items" array
    let fetchAndDeserializeLedgersAsync requestLimit encodedCredentials balancesDisplayFields balancesFinder () =
        let rec helper (offset: int) (listAcc: BalancesFields list) : Async<BalancesFields list> =
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
                        | Some deserializedResponse ->
                            let fetchedItems = deserializedResponse.Items
                            let hasMore = deserializedResponse.HasMore
                            //printfn "Successfully fetched and deserialized %d balances." (List.length fetchedItems)
                            match hasMore with
                            | true -> return! helper (offset + requestLimit) (listAcc @ fetchedItems)
                            | _ -> return listAcc @ fetchedItems
                        | None ->
                            // Deserialization failed; return the accumulated list so far
                            printfn "Failed to deserialize response."
                            errMSG <- "Failed to deserialize response."
                            return listAcc
                    | HttpStatusCode.Unauthorized ->
                        printfn "Unauthorized: Check your credentials."
                        errMSG <- "Unauthorized: Check your credentials."
                        return listAcc
                    | HttpStatusCode.Forbidden ->
                        printfn "Forbidden: You don't have access to this resource."
                        errMSG <- "Forbidden: You don't have access to this resource."
                        return listAcc
                    | HttpStatusCode.NotFound ->
                        printfn "Not Found: The requested resource does not exist."
                        errMSG <- "Not Found: The requested resource does not exist."
                        return listAcc
                    | status ->
                        printfn "Request failed with status code %d" (int status)

                        errMSG <-
                            "Request failed with status code: "
                            + status.ToString()

                        return listAcc
                with
                | ex ->
                    printfn "An error occurred: %s" ex.Message
                    errMSG <- "An error occurred: " + ex.Message
                    return listAcc
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
        try
            // Define your credentials and requestLimit
            let requestLimit = 500

            let encodedCredentials =
                (GetSecret "oracleuser", GetSecret "oraclepassword")
                |> encodeBasicAuth
            // Fetch and deserialize balances asynchronously
            let data =
                fetchAndDeserializeLedgersAsync requestLimit encodedCredentials balancesDisplayFields balancesFinder ()
                |> Async.RunSynchronously
                |> List.map balancesFieldsToList

            if List.isEmpty data then
                // Return a message indicating no data was fetched
                errMSG
            else
                // Extract headers from the first record to preserve order
                let headers =
                    match data with
                    | firstRecord :: _ -> firstRecord |> List.map fst
                    | [] -> []

                // Determine the maximum number of segments across all records
                // Exclude records where DetailAccountCombination is 'N/A'
                let maxSegments =
                    data
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
                    data
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
                            | Some (:? string as s) when s <> "N/A" -> splitDetailAccountCombination s |> Array.map box
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
                box resultArray
        with
        | ex ->
            // In case of any unexpected errors, return the error message in Excel
            ("An error occurred: " + ex.Message) :> obj
