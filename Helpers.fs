module OraGlBalEx.Helpers
    open System
    open System.Text
    
    open JsonConverters

    /// Encodes username and password for Basic Authentication
    let encodeBasicAuth (username: string, password: string) : string =
        let credentials = $"{username}:{password}"
        let bytes = Encoding.UTF8.GetBytes(credentials)
        Convert.ToBase64String(bytes)


    /// Extracts a float value from FloatOrString, returning a default if it's a string
    let getFloatOrDefault (fos: FloatOrString) (defaultValue: float) : float =
        match fos with
        | FloatOrString.FloatValue num -> num
        | FloatOrString.StringValue _ -> defaultValue

    /// Checks if a FloatOrString is "N/A"
    let isNA (fos: FloatOrString) : bool =
        match fos with
        | FloatOrString.StringValue str when str = "N/A" -> true
        | _ -> false

    /// Converts FloatOrString to float option
    let toFloatOption (fos: FloatOrString) : float option =
        match fos with
        | FloatOrString.FloatValue num -> Some num
        | FloatOrString.StringValue _ -> None
