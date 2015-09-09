defmodule Mix.Tasks.Step6File do
  import Mal.Types
  alias Mal.Function

  def run(args) do
    env = Mal.Env.initialize()
    Mal.Env.merge(env, Mal.Core.namespace)
    bootstrap(args, env)
    load_file(args, env)
    loop(env)
  end

  defp load_file([], _env), do: nil
  defp load_file([file_name | _args], env) do
    read_eval_print("""
      (load-file "#{file_name}")
      """, env)
    exit(:normal)
  end

  defp bootstrap(args, env) do
    # not:
    read_eval_print("""
      (def! not
        (fn* (a) (if a false true)))
      """, env)

    # load-file:
    read_eval_print("""
      (def! load-file
        (fn* (f)
          (eval (read-string (str "(do " (slurp f) ")")))))
      """, env)

    Mal.Env.set(env, "eval", %Function{value: fn [ast] ->
      eval(ast, env)
    end})

    case args do
      [_file_name | rest] -> Mal.Env.set(env, "*ARGV*", list(rest))
      [] -> Mal.Env.set(env, "*ARGV*", list([]))
    end
  end

  defp loop(env) do
    IO.write(:stdio, "user> ")
    IO.read(:stdio, :line)
      |> read_eval_print(env)
      |> IO.puts

    loop(env)
  end

  defp eval_ast({:list, ast, meta}, env) when is_list(ast) do
    {:list, Enum.map(ast, fn elem -> eval(elem, env) end), meta}
  end

  defp eval_ast({:map, ast, meta}, env) do
    map = for {key, value} <- ast, into: %{} do
      {eval(key, env), eval(value, env)}
    end

    {:map, map, meta}
  end

  defp eval_ast({:vector, ast, meta}, env) do
    {:vector, Enum.map(ast, fn elem -> eval(elem, env) end), meta}
  end

  defp eval_ast({:symbol, symbol}, env) do
    case Mal.Env.get(env, symbol) do
      {:ok, value} -> value
      :not_found -> throw({:error, "'#{symbol}' not found"})
    end
  end

  defp eval_ast(ast, _env), do: ast

  defp read(input) do
    Mal.Reader.read_str(input)
  end

  defp eval_bindings([], _env), do: _env
  defp eval_bindings([{:symbol, key}, binding | tail], env) do
    evaluated = eval(binding, env)
    Mal.Env.set(env, key, evaluated)
    eval_bindings(tail, env)
  end
  defp eval_bindings(_bindings, _env), do: throw({:error, "Unbalanced let* bindings"})

  defp eval({:list, ast, meta}, env), do: eval_list(ast, env, meta)
  defp eval(ast, env), do: eval_ast(ast, env)

  defp eval_list([{:symbol, "if"}, condition, if_true | if_false], env, _) do
    result = eval(condition, env)
    if result == nil or result == false do
      case if_false do
        [] -> nil
        [body] -> eval(body, env)
      end
    else
      eval(if_true, env)
    end
  end

  defp eval_list([{:symbol, "do"} | ast], env, _) do
    ast
      |> List.delete_at(-1)
      |> list
      |> eval_ast(env)
    eval(List.last(ast), env)
  end

  defp eval_list([{:symbol, "def!"}, {:symbol, key}, value], env, _) do
    evaluated = eval(value, env)
    Mal.Env.set(env, key, evaluated)
    evaluated
  end

  defp eval_list([{:symbol, "let*"}, {list_type, bindings, _}, body], env, _)
  when list_type == :list or list_type == :vector do
    let_env = Mal.Env.initialize(env)
    eval_bindings(bindings, let_env)
    eval(body, let_env)
  end

  defp eval_list([{:symbol, "fn*"}, {list_type, params, _}, body], env, _)
  when list_type == :list or list_type == :vector do
    param_symbols = for {:symbol, symbol} <- params, do: symbol

    closure = fn args ->
      inner = Mal.Env.initialize(env, param_symbols, args)
      eval(body, inner)
    end

    %Function{value: closure}
  end

  defp eval_list(ast, env, meta) do
    {:list, [func | args], _} = eval_ast({:list, ast, meta}, env)
    func.value.(args)
  end

  defp print(value) do
    Mal.Printer.print_str(value)
  end

  defp read_eval_print(:eof, _env), do: exit(:normal)
  defp read_eval_print(line, env) do
    read(line)
      |> eval(env)
      |> print
  catch
    {:error, message} -> IO.puts("Error: #{message}")
  end
end
