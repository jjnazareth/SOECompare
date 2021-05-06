module ExcelData

open System
open System.IO
open ExcelDataReader

type Status =
    | Ok
    | Error

type ExcelRow =
    { Diocese: string
      Title: string
      FirstName: string
      LastName: string
      EmailAddress: string
      IsEPrayer: Boolean
      Status: Status }

type ExcelAllRow =
    { Diocese: string
      Title: string
      FirstName: string
      LastName: string
      Address: string
      Telephone: string
      Mobile: string
      EmailAddress: string
      IsEPrayer: Boolean
      Status: Status }


let getExcelData path (sheet: string) =
    use stream =
        File.Open(path, FileMode.Open, FileAccess.Read)

    let reader =
        ExcelReaderFactory.CreateOpenXmlReader(stream)

    let getData (reader: IExcelDataReader) =
        let convertToDouble (o: obj) =
            match o with
            | :? Double as x -> x
            | _ -> 0.0

        let tables = reader.AsDataSet().Tables
        let table = tables.get_Item (sheet)
        seq {
            for row in table.Rows ->
                { Diocese = Convert.ToString row.[0]
                  Title = Convert.ToString row.[1]
                  FirstName = Convert.ToString row.[2]
                  LastName = Convert.ToString row.[3]
                  EmailAddress = Convert.ToString row.[10]
                  IsEPrayer = (Convert.ToString row.[13]).Trim().ToUpper() = "E" // row.[12] earlier
                  Status = if Convert.ToString row.[11] = "" then Ok else Error }
        }

    getData reader
    |> Seq.skip 1 // header
    |> Seq.filter (fun x -> x.EmailAddress <> "")

let getExcelAllData path (sheet: string) =
    use stream =
        File.Open(path, FileMode.Open, FileAccess.Read)

    let reader =
        ExcelReaderFactory.CreateOpenXmlReader(stream)

    let getData (reader: IExcelDataReader) =
        let convertToDouble (o: obj) =
            match o with
            | :? Double as x -> x
            | _ -> 0.0

        let tables = reader.AsDataSet().Tables
        let table = tables.get_Item (sheet)
        seq {
            for row in table.Rows ->
                { Diocese = Convert.ToString row.[0]
                  Title = Convert.ToString row.[1]
                  FirstName = Convert.ToString row.[2]
                  LastName = Convert.ToString row.[3]
                  Address =
                      [ 4; 5; 6; 7 ]
                      |> List.map (fun x -> Convert.ToString row.[x])
                      |> List.reduce (fun acc x -> acc + " " + x)
                  Telephone = Convert.ToString row.[8]
                  Mobile = Convert.ToString row.[9]
                  EmailAddress = Convert.ToString row.[10]
                  IsEPrayer = (Convert.ToString row.[13]).Trim().ToUpper() = "E" // row.[12] earlier
                  Status = if Convert.ToString row.[11] = "" then Ok else Error }
        }

    getData reader |> Seq.skip 1 // header
