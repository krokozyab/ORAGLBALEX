module OraGlBalEx.Security

open System
open ExcelDna.Integration
open CredentialManagement
open Serilog

/// Stores a secret in the Windows Credential Manager.
/// Accepts a single string input in the format "key: value".
[<ExcelFunction(Description = "Stores a secret in Windows Credential Manager. Input format: \"key: value\".")>]
let StoreSecret (input: string) : string =
    if String.IsNullOrWhiteSpace(input) then
        "#ERROR: Input cannot be empty."
    else
        // Split the input into key and value
        let parts = input.Split([| ':' |], 2, StringSplitOptions.RemoveEmptyEntries)

        if parts.Length <> 2 then
            "#ERROR: Input must be in the format \"key: value\"."
        else
            let key = parts[ 0 ].Trim()
            let value = parts[ 1 ].Trim()

            if
                String.IsNullOrWhiteSpace(key)
                || String.IsNullOrWhiteSpace(value)
            then
                "#ERROR: Key and value cannot be empty."
            else
                try
                    // Create and set the credential
                    let cred = new Credential()
                    cred.Target <- key
                    cred.Username <- key // You can modify this as needed
                    cred.Password <- value
                    cred.Type <- CredentialType.Generic
                    cred.PersistanceType <- PersistanceType.LocalComputer // Or .Session for session-only
                    let success = cred.Save()

                    if success then
                        sprintf "Secret for '%s' stored successfully." key
                    else
                        Log.Error("#ERROR: Failed to store the secret.")
                        "#ERROR: Failed to store the secret."
                with
                | ex ->
                    Log.Error($"#ERROR: {ex.Message}")
                    sprintf $"#ERROR: {ex.Message}"

/// Retrieves a secret from the Windows Credential Manager by key.
/// Not accessible from Excel ui
//[<ExcelFunction(Description = "Retrieves a secret from Windows Credential Manager by key.")>]
let GetSecret
    ([<ExcelArgument(Name = "Sectet name", Description = "Secret name.", AllowReference = true)>] key: string)
    : string =
    if String.IsNullOrWhiteSpace(key) then
        "#ERROR: Key cannot be empty."
    else
        try
            // Initialize the credential
            let cred = new Credential(Target = key, Type = CredentialType.Generic)
            let loaded = cred.Load()

            if loaded then
                cred.Password
            else
                "#ERROR: No secret found for the given key."
        with
        | ex -> sprintf "#ERROR: %s" ex.Message

//Deletes a secret from Windows Credential Manager by key
[<ExcelFunction(Description = "Deletes a secret from Windows Credential Manager by key.")>]
let DeleteSecret (key: string) : string =
    if String.IsNullOrWhiteSpace(key) then
        "#ERROR: Key cannot be empty."
    else
        try
            // Initialize the credential
            let cred = new Credential(Target = key, Type = CredentialType.Generic)
            let loaded = cred.Load()

            if loaded then
                let success = cred.Delete()

                if success then
                    sprintf "Secret for '%s' deleted successfully." key
                else
                    Log.Error("#ERROR: Failed to delete the secret: " + key)
                    "#ERROR: Failed to delete the secret: " + key
            else
                "#ERROR: No secret found for the given key."
        with
        | ex -> sprintf "#ERROR: %s" ex.Message
