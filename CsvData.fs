module CsvData

open System.IO
open FSharp.Data

type Status =
    | Ok
    | Error

type ExcelRow =
    { Diocese: string
      Title: string
      FirstName: string
      LastName: string
      EmailAddress: string
      Status: Status }


let getCsvData (path: string) =
    let csv = CsvFile.Load(path, ",", '"', true, true)
    csv.Rows
    |> Seq.map (fun row ->
        { EmailAddress = row.[0]
          FirstName = row.[1]
          LastName = row.[2]
          Title = row.[3]
          Diocese = row.[4]
          Status = Ok })
