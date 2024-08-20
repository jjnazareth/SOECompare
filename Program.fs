
open System.IO
open ExcelData
open CsvData
open FSharp.Collections

let tee f x = f x; x

let folder = @"C:\Users\Jivraj\Documents\Jivraj\SOE\Db"
let fileName = @"Full Dbase Open 2024-07-03.xlsx"
let sheet = "Master"
let xlFPath = Path.Combine (folder, fileName)
    
let clnFName = @"cleaned_email_segment_export_a4f09ca51a.csv"
let unsubFName = @"unsubscribed_email_segment_export_a4f09ca51a.csv"
let subFName = @"subscribed_email_segment_export_a0afb966c3.csv"
let ePrayerFName =  @"eprayers_segment_export_41058b5d9b.csv"

let noDbSubscribed =
    // subscribed in MailChimp but not in Full Dbase.xlsx
    let dbEmails =
        getExcelData xlFPath sheet
        |> Seq.map (fun x -> x.EmailAddress.Trim().ToLower())
        |> Seq.toList

    let subEmails =
        let path = Path.Combine (folder, subFName)
        query {
           for data in getCsvData path do
           select data.EmailAddress
        }
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList
        // |> tee (List.iter (printfn "%s"))
    query {
        for email in subEmails do
        where (not <| List.contains email dbEmails )
    }

let dbUnsubscribed =
    // exist in Full Dbase.xlsx but are unsubscribed
    let path = Path.Combine (folder, unsubFName)
    let unSubEmails  = 
        query {
            for data in getCsvData path do
            select data.EmailAddress
            } 
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    let dbEmails =
        getExcelData xlFPath sheet
        |> Seq.map (fun x -> x.EmailAddress.Trim().ToLower())
        |> Seq.toList

    query {
        for email in dbEmails do 
        where(List.contains email unSubEmails)
        select email
    }

let dbCleaned =
    // cleaned in MailChimp but showing in Full Dbase.xlsx
    let path = Path.Combine (folder, clnFName)
    let clnEmails  = 
        query {
            for data in getCsvData path do
            select data.EmailAddress
            } 
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    let dbEmails =
        getExcelData xlFPath sheet
        |> Seq.map (fun x -> x.EmailAddress.Trim().ToLower())
        |> Seq.toList

    query {
        for email in dbEmails do 
        where(List.contains email clnEmails)
        select email
    }

let dbNotMailChimp =
    // existing in Full Dbase.xlsx but not subscribed or unsubscribed or cleaned in MailChimp
    let subEmails  = 
        let subPath = Path.Combine (folder, subFName)
        query {
            for data in getCsvData subPath do
            select data.EmailAddress
            } 
            
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    let unSubEmails  = 
        let unsubPath = Path.Combine (folder, unsubFName)
        query {
            for data in getCsvData unsubPath do
            select data.EmailAddress
            } 
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    let path = Path.Combine (folder, clnFName)

    let clnEmails  = 
        query {
            for data in getCsvData path do
            select data.EmailAddress
            } 
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    let dbEmails =
        getExcelData xlFPath sheet
        |> Seq.map (fun x -> x.EmailAddress.Trim().ToLower())
        |> Seq.toList

    query {
        for email in dbEmails do 
        where ( (not <| List.contains email subEmails) &&
                (not <| List.contains email unSubEmails) &&  
                (not <| List.contains email clnEmails)  )
        select email
    }

let noDbEPrayers =
    let dbEPrayers =
        getExcelData xlFPath sheet
        |> Seq.filter (fun x -> x.IsEPrayer)
        |> Seq.map (fun x -> x.EmailAddress.Trim().ToLower())
        |> Seq.toList
    let ePrayerEmails =
        let path = Path.Combine (folder, ePrayerFName)
        query {
           for data in getCsvData path do
           select data.EmailAddress
        }
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList
    query {
        for email in ePrayerEmails do
        where (not <| List.contains email dbEPrayers )
    }

let dbNotMailChimpEPrayers =
    // existing in Full Dbase.xlsx but not in MailChimp EPrayers
    let ePrayerEmails =
        let path = Path.Combine (folder, ePrayerFName)
        query {
           for data in getCsvData path do
           select data.EmailAddress
        }
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    let dbEPrayers =
        getExcelData xlFPath sheet
        |> Seq.filter (fun x -> x.IsEPrayer)
        |> Seq.map (fun x -> x.EmailAddress.Trim().ToLower())
        |> Seq.toList

    query {
        for email in dbEPrayers do 
        where (not <| List.contains email ePrayerEmails)
        select email
    }

let newContactsCsv =
    // existing in Full Dbase.xlsx but not subscribed or unsubscribed in MailChimp
    let subEmails  = 
        let subPath = Path.Combine (folder, subFName)
        query {
            for data in getCsvData subPath do
            select data.EmailAddress
            } 
            
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    let unSubEmails  = 
        let unsubPath = Path.Combine (folder, unsubFName)
        query {
            for data in getCsvData unsubPath do
            select data.EmailAddress
            } 
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    query {
        for row in getExcelData xlFPath sheet do 
        where ( (not <| List.contains (row.EmailAddress.Trim().ToLower()) subEmails) && 
                (not <| List.contains (row.EmailAddress.Trim().ToLower()) unSubEmails )  )
            
        select row
    }
     
let existingContactsCsv =
    // existing in Full Dbase.xlsx and not unsubscribed in MailChimp

    let unSubEmails  = 
        let unsubPath = Path.Combine (folder, unsubFName)
        query {
            for data in getCsvData unsubPath do
            select data.EmailAddress
            } 
        |> Seq.map (fun x -> x.Trim().ToLower())
        |> Seq.toList

    query {
        for row in getExcelData xlFPath sheet do 
        where  (not <| List.contains (row.EmailAddress.Trim().ToLower()) unSubEmails )            
        select row
    }
     

[<EntryPoint>]
let main argv = 

    if argv.Length = 0 then 
        printfn "Records that do not exist in Full Dbase.xlsx but are subscribed"
        noDbSubscribed
        |> Seq.iter (printfn "%s")
        printfn "==============="
        printfn "Records that exist in Full Dbase.xlsx but are unsubscribed"
        dbUnsubscribed
        |> Seq.iter (printfn "%s")
        printfn "==============="
        printfn "Records that exist in Full Dbase.xlsx but are not in MailChimp"
        dbNotMailChimp
        |> Seq.iter (printfn "%s")
        printfn "==============="
        printfn "Records that exist in Full Dbase.xlsx but are cleaned in MailChimp"
        dbCleaned
        |> Seq.iter (printfn "%s")
        printfn "==============="
        printfn "Records that are not in ePrayers in Full Dbase.xlsx but are ePrayers in MailChimp"
        noDbEPrayers
        |> Seq.iter (printfn "%s")    
        printfn "==============="

        printfn "Records that exist in Full Dbase.xlsx but are not in ePrayers in MailChimp"
        dbNotMailChimpEPrayers
        |> Seq.iter (printfn "%s")    
        printfn "==============="
    else 
        printfn "Diocese, Email Address, Title, First Name, Last Name"
        newContactsCsv
        |> Seq.iter (fun x -> printfn "%s,%s,%s,%s,%s" x.Diocese (x.EmailAddress.Trim()) x.Title x.FirstName x.LastName )
 

    0 // return an integer exit code