module OraGlBalEx.JsonConverters

open System
open Newtonsoft.Json
open Newtonsoft.Json.Linq

/// Represents a value that can be either a float or a string
type FloatOrString =
    | FloatValue of float
    | StringValue of string

type FloatOrStringConverter() =
    inherit JsonConverter()

    override _.CanConvert(objectType: Type) = objectType = typeof<FloatOrString>

    override _.ReadJson(reader: JsonReader, objectType: Type, existingValue: obj, serializer: JsonSerializer) : obj =
        match reader.TokenType with
        | JsonToken.Float
        | JsonToken.Integer ->
            let num = Convert.ToDouble(reader.Value)
            FloatOrString.FloatValue(num) :> obj
        | JsonToken.String ->
            let str = reader.Value :?> string

            if str = "N/A" then
                FloatOrString.StringValue(str) :> obj
            else
                match
                    Double.TryParse
                        (
                            str,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture
                        )
                    with
                | (true, num) -> FloatOrString.FloatValue(num) :> obj
                | _ -> FloatOrString.StringValue(str) :> obj
        | JsonToken.Null ->
            // Treat null as "N/A" or handle accordingly
            FloatOrString.StringValue("N/A") :> obj
        | _ ->
            // For any other unexpected token, throw an exception
            raise (JsonSerializationException("Unexpected token type for FloatOrString."))

    override _.WriteJson(writer: JsonWriter, value: obj, serializer: JsonSerializer) =
        match value with
        | :? FloatOrString as fos ->
            match fos with
            | FloatOrString.FloatValue num -> writer.WriteValue(num)
            | FloatOrString.StringValue str -> writer.WriteValue(str)
        | _ -> writer.WriteNull()

type OptionConverter() =
    inherit JsonConverter()

    override _.CanConvert(objectType: Type) =
        objectType.IsGenericType
        && objectType.GetGenericTypeDefinition() = typedefof<option<_>>

    override _.ReadJson(reader: JsonReader, objectType: Type, existingValue: obj, serializer: JsonSerializer) : obj =
        let innerType = objectType.GetGenericArguments().[0]
        let jToken = JToken.Load(reader)

        if jToken.Type = JTokenType.Null then
            // Return None
            null
        else
            // Deserialize the inner value and wrap it in Some
            let value = jToken.ToObject(innerType, serializer)
            // Create Some(value)
            let someCase =
                FSharp.Reflection.FSharpValue.MakeUnion(
                    FSharp.Reflection.FSharpType.GetUnionCases(objectType)
                    |> Array.find (fun uc -> uc.Name = "Some"),
                    [| value |]
                )

            someCase

    override _.WriteJson(writer: JsonWriter, value: obj, serializer: JsonSerializer) =
        match value with
        | null -> writer.WriteNull()
        | _ ->
            let unionCaseInfo, fields =
                FSharp.Reflection.FSharpValue.GetUnionFields(value, value.GetType())

            match unionCaseInfo.Name with
            | "Some" ->
                // Serialize the inner value
                serializer.Serialize(writer, fields[0])
            | "None" -> writer.WriteNull()
            | _ -> raise (JsonSerializationException(sprintf "Unknown union case: %s" unionCaseInfo.Name))

/// Configure JsonSerializerSettings to include both custom converters
let jsonSettings =
    let settings = JsonSerializerSettings()
    // Register the OptionConverter first
    settings.Converters.Add(OptionConverter())
    // Register the FloatOrStringConverter
    settings.Converters.Add(FloatOrStringConverter())
    // Additional settings
    settings.NullValueHandling <- NullValueHandling.Ignore
    settings.MissingMemberHandling <- MissingMemberHandling.Ignore
    settings
