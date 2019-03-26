Attribute VB_Name = "LIntFunDeclare"
Option Explicit

Public Const internalDefs = "(do " & _
"(def! load-file (fn* (f) (eval (read-string (str ""(do "" (slurp f) "")""))))) " & _
"(defmacro! cond (fn* (& xs) (if (> (count xs) 0) (list 'if (first xs) (if (> (count xs) 1) (nth xs 1) (throw ""odd number of forms to cond"")) (cons 'cond (rest (rest xs))))))) " & _
"(def! *gensym-counter* (atom 0))" & _
"(def! gensym (fn* [] (symbol (str ""G__"" (swap! *gensym-counter* (fn* [x] (+ 1 x)))))))" & _
"(defmacro! or (fn* (& xs) (if (empty? xs) nil (if (= 1 (count xs)) (first xs) (let* (condvar (gensym)) `(let* (~condvar ~(first xs)) (if ~condvar ~condvar (or ~@(rest xs)))))))))" & _
")"
    
    
Sub declareInternalFunctions(env As LHashmap)
    env.add "+", newInternalFunction(add)
    env.add "-", newInternalFunction(subtract)
    env.add "*", newInternalFunction(multiply)
    env.add "/", newInternalFunction(divide)
    env.add "pr-str", newInternalFunction(pr_str_)
    env.add "str", newInternalFunction(str_)
    env.add "list", newInternalFunction(list)
    env.add "list?", newInternalFunction(list_q)
    env.add "empty?", newInternalFunction(empty_q)
    env.add "count", newInternalFunction(count_)
    env.add "=", newInternalFunction(eq)
    env.add "<", newInternalFunction(lt)
    env.add "<=", newInternalFunction(le)
    env.add ">", newInternalFunction(gt)
    env.add ">=", newInternalFunction(ge)
    env.add "not", newInternalFunction(not_)
    env.add "read-string", newInternalFunction(read_string)
    env.add "slurp", newInternalFunction(slurp)
    
    env.add "atom", newInternalFunction(lif_atom)
    env.add "atom?", newInternalFunction(lif_atom_q)
    env.add "deref", newInternalFunction(lif_deref)
    env.add "reset!", newInternalFunction(lif_reset_)
    env.add "swap!", newInternalFunction(lif_swap_)
    
    env.add "cons", newInternalFunction(lif_cons)
    env.add "concat", newInternalFunction(lif_concat)
    env.add "nth", newInternalFunction(lif_nth)
    env.add "slice", newInternalFunction(lif_slice)
    env.add "str?", newInternalFunction(lif_str_p)
    env.add "charcodeat", newInternalFunction(lif_charcodeat)
    env.add "first", newInternalFunction(lif_first)
    env.add "rest", newInternalFunction(lif_rest)
    
    env.add "apply", newInternalFunction(lif_apply)
    env.add "map", newInternalFunction(lif_map)
    env.add "throw", newInternalFunction(lif_throw)
    
    env.add "nil?", newInternalFunction(lif_nil_p)
    env.add "true?", newInternalFunction(lif_true_p)
    env.add "false?", newInternalFunction(lif_false_p)
    env.add "symbol?", newInternalFunction(lif_symbol_p)
    
    env.add "symbol", newInternalFunction(lif_symbol)
    env.add "keyword", newInternalFunction(lif_keyword)
    env.add "keyword?", newInternalFunction(lif_keyword_p)
    env.add "vector", newInternalFunction(lif_vector)
    env.add "vector?", newInternalFunction(lif_vector_p)
    env.add "hash-map", newInternalFunction(lif_hash_map)
    env.add "map?", newInternalFunction(lif_map_p)
    env.add "assoc", newInternalFunction(lif_assoc)
    env.add "dissoc", newInternalFunction(lif_dissoc)
    env.add "get", newInternalFunction(lif_get)
    env.add "contains?", newInternalFunction(lif_contains_p)
    env.add "keys", newInternalFunction(lif_keys)
    env.add "vals", newInternalFunction(lif_vals)
    env.add "sequential?", newInternalFunction(lif_sequential_p)
    
    env.add "callbyname", newInternalFunction(lif_callbyname)
    env.add "time-ms", newInternalFunction(lif_time_ms)
    
End Sub
