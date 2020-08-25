namespace VMS.TPS

open System.Diagnostics
open System.Reflection
open System.Windows
open VMS.TPS.Common.Model.API
open VMS.TPS.Common.Model.Types

type Script () =

    member this.Execute (context:ScriptContext) =

        let fullAssemblyName = Assembly.GetExecutingAssembly().Location
        let minusExtension = fullAssemblyName.[0..fullAssemblyName.Length-11]
        let exePath = minusExtension + ".exe"

        try
            Process.Start(exePath) |> ignore
        with ex ->
            MessageBox.Show(ex.ToString(), "Failed to start application.") |> ignore
