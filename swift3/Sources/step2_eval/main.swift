import Foundation

// read
func READ(str: String) throws -> MalVal {
    return try read_str(str)
}

// eval
func eval_ast(ast: MalVal, _ env: Dictionary<String, MalVal>) throws -> MalVal {
    switch ast {
    case MalVal.MalSymbol(let sym):
        if env[sym] == nil {
            throw MalError.General(msg: "'\(sym)' not found")
        }
        return env[sym]!
    case MalVal.MalList(let lst):
        return MalVal.MalList(try lst.map { try EVAL($0, env) })
    case MalVal.MalVector(let lst):
        return MalVal.MalVector(try lst.map { try EVAL($0, env) })
    case MalVal.MalHashMap(let dict):
        var new_dict = Dictionary<String,MalVal>()
        for (k,v) in dict { new_dict[k] = try EVAL(v, env) }
        return MalVal.MalHashMap(new_dict)
    default:
        return ast
    }
}

func EVAL(ast: MalVal, _ env: Dictionary<String, MalVal>) throws -> MalVal {
    switch ast {
    case MalVal.MalList: true
    default: return try eval_ast(ast, env)
    }

    switch try eval_ast(ast, env) {
    case MalVal.MalList(let elst):
        switch elst[0] {
        case MalVal.MalFunc(let fn,_,_,_,_,_):
            let args = Array(elst[1..<elst.count])
            return try fn(args)
        default:
            throw MalError.General(msg: "Cannot apply on '\(elst[0])'")
        }
    default: throw MalError.General(msg: "Invalid apply")
    }
}

// print
func PRINT(exp: MalVal) -> String {
    return pr_str(exp, true)
}


// repl
func rep(str:String) throws -> String {
    return PRINT(try EVAL(try READ(str), repl_env))
}

func IntOp(op: (Int, Int) -> Int, _ a: MalVal, _ b: MalVal) throws -> MalVal {
    switch (a, b) {
    case (MalVal.MalInt(let i1), MalVal.MalInt(let i2)):
        return MalVal.MalInt(op(i1, i2))
    default:
        throw MalError.General(msg: "Invalid IntOp call")
    }
}

var repl_env: Dictionary<String,MalVal> = [
    "+": MalVal.MalFunc({ try IntOp({ $0 + $1}, $0[0], $0[1]) },
                                ast:nil, env:nil, params:nil,
                                macro:false, meta:nil),
    "-": MalVal.MalFunc({ try IntOp({ $0 - $1}, $0[0], $0[1]) },
                                ast:nil, env:nil, params:nil,
                                macro:false, meta:nil),
    "*": MalVal.MalFunc({ try IntOp({ $0 * $1}, $0[0], $0[1]) },
                                ast:nil, env:nil, params:nil,
                                macro:false, meta:nil),
    "/": MalVal.MalFunc({ try IntOp({ $0 / $1}, $0[0], $0[1]) },
                                ast:nil, env:nil, params:nil,
                                macro:false, meta:nil),
]

while true {
    print("user> ", terminator: "")
    let line = readLine(stripNewline: true)
    if line == nil { break }
    if line == "" { continue }

    do {
        print(try rep(line!))
    } catch (MalError.Reader(let msg)) {
        print("Error: \(msg)")
    } catch (MalError.General(let msg)) {
        print("Error: \(msg)")
    }
}
