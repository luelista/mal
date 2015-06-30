module REPL
    open System
    open Types

    let rec eval_ast env = function
        | Symbol(sym) -> Env.get env sym
        | List(lst) -> lst |> List.map (eval env) |> List
        | Vector(seg) -> seg |> Seq.map (eval env) |> Array.ofSeq |> Node.ofArray
        | Map(map) -> map |> Map.map (fun k v -> eval env v) |> Map
        | node -> node

    and eval env = function
        | List(_) as node ->
            let resolved = node |> eval_ast env
            match resolved with
            | List(Func(_, f, _, _, [])::rest) -> f rest
            | _ -> raise <| Error.errExpectedX "function"
        | node -> node |> eval_ast env

    let READ input =
        try
            Reader.read_str input
        with
        | Error.ReaderError(msg) ->
            printfn "%s" msg
            []

    let EVAL env ast =
        try
            Some(eval env ast)
        with
        | Error.EvalError(msg) 
        | Error.ReaderError(msg) ->
            printfn "%s" msg
            None

    let PRINT v =
        v
        |> Seq.singleton
        |> Printer.pr_str
        |> printfn "%s"

    let REP env input =
        READ input
        |> Seq.ofList
        |> Seq.choose (fun form -> EVAL env form)
        |> Seq.iter (fun value -> PRINT value)

    let getReadlineMode (args : string array) =
        if args.Length > 0 && args.[0] = "--raw" then
            Readline.Mode.Raw
        else
            Readline.Mode.Terminal

    [<EntryPoint>]
    let rec main args =
        let mode = getReadlineMode args
        let env = Env.makeRootEnv ()
        match Readline.read "user> " mode with
        | null -> 0
        | input -> 
            REP env input
            main args
