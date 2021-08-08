// Learn more about F# at http://docs.microsoft.com/dotnet/fsharp

open System
open Midoliy.Office.Interop

// Define a function to construct a message to print
let fn () =
  use app = Excel.BlankWorkbook()
  app.Visibility <- AppVisibility.Visible
  let sheet = app.[1].[1]
  sheet.["A1:A3"].Value <- 100
  sheet.["A1:A3"].Select()
  sheet.Shapes.AddChart(ChartRecipe.MakeLine()) |> ignore
  ()

[<EntryPoint>]
let main argv =
  fn()
  0 // return an integer exit code

