defmodule Mal.Core do
  import Mal.Types
  alias Mal.Function

  def namespace do
    %{
      "+" => fn [a, b] -> a + b end,
      "-" => fn [a, b] -> a - b end,
      "*" => fn [a, b] -> a * b end,
      "/" => fn [a, b] -> div(a, b) end,
      ">" => fn [a, b] -> a > b end,
      "<" => fn [a, b] -> a < b end,
      "<=" => fn [a, b] -> a <= b end,
      ">=" => fn [a, b] -> a >= b end,
      "concat" => &concat/1,
      "=" => &equal/1,
      "list?" => &list?/1,
      "empty?" => &empty?/1,
      "count" => &count/1,
      "pr-str" => &pr_str/1,
      "str" => &str/1,
      "prn" => &prn/1,
      "println" => &println/1,
      "slurp" => &slurp/1,
      "nth" => &nth/1,
      "first" => &first/1,
      "rest" => &rest/1,
      "map" => &map/1,
      "apply" => &apply/1,
      "keyword" => &keyword/1,
      "symbol?" => &symbol?/1,
      "cons" => &cons/1,
      "vector?" => &vector?/1,
      "assoc" => &assoc/1,
      "dissoc" => &dissoc/1,
      "get" => &get/1,
      "map?" => &map?/1,
      "list" => &list/1,
      "vector" => &vector/1,
      "hash-map" => &hash_map/1,
      "readline" => fn [prompt] -> readline(prompt) end,
      "sequential?" => fn arg -> vector?(arg) or list?(arg) end,
      "keyword?" => fn [type] -> is_atom(type) end,
      "nil?" => fn [type] -> type == nil end,
      "true?" => fn [type] -> type == true end,
      "false?" => fn [type] -> type == false end,
      "symbol" => fn [name] -> {:symbol, name} end,
      "read-string" => fn [input] -> Mal.Reader.read_str(input) end,
      "throw" => fn [arg] -> throw({:error, arg}) end,
      "contains?" => fn [{:map, map, _}, key] -> Map.has_key?(map, key) end,
      "keys" => fn [{:map, map, _}] -> Map.keys(map) |> list end,
      "vals" => fn [{:map, map, _}] -> Map.values(map) |> list end
    }
  end

  def readline(prompt) do
    IO.write(:stdio, prompt)
    IO.read(:stdio, :line)
      |> String.strip(?\n)
  end

  defp convert_vector({:vector, ast, meta}), do: {:list, ast, meta}
  defp convert_vector(other), do: other

  defp equal([a, b]) do
    convert_vector(a) == convert_vector(b)
  end

  defp list?([{:list, _, _}]), do: true
  defp list?(_), do: false

  defp empty?([{_type, [], _meta}]), do: true
  defp empty?(_), do: false

  defp count([{_type, ast, _meta}]), do: length(ast)
  defp count(_), do: 0

  defp pr_str(args) do
    args
      |> Enum.map(&Mal.Printer.print_str/1)
      |> Enum.join(" ")
  end

  defp str(args) do
    args
      |> Enum.map(&(Mal.Printer.print_str(&1, false)))
      |> Enum.join("")
  end

  defp prn(args) do
    args
      |> pr_str
      |> IO.puts
    nil
  end

  defp println(args) do
    args
      |> Enum.map(&(Mal.Printer.print_str(&1, false)))
      |> Enum.join(" ")
      |> IO.puts
    nil
  end

  defp slurp([file_name]) do
    case File.read(file_name) do
      {:ok, content} -> content
      {:error, :enoent} -> throw({:error, "can't find file #{file_name}"})
      {:error, :eisdir} -> throw({:error, "can't read directory #{file_name}"})
      {:error, :eaccess} -> throw({:error, "missing permissions #{file_name}"})
      {:error, reason} -> throw({:error, "can't read file #{file_name}, #{reason}"})
    end
  end

  defp nth([{_type, ast, _meta}, index]) do
    case Enum.at(ast, index, :error) do
      :error -> throw({:error, "index out of bounds"})
      any -> any
    end
  end

  defp first([{_type, [head | tail], _}]), do: head
  defp first(_), do: nil

  defp rest([{_type, [head | tail], _}]), do: list(tail)
  defp rest([{_type, [], _}]), do: list([])

  defp map([%Function{value: function}, ast]), do: do_map(function, ast)
  defp map([function, ast]), do: do_map(function, ast)

  defp do_map(function, {_type, ast, _meta}) do
    ast
      |> Enum.map(fn arg -> function.([arg]) end)
      |> list
  end

  defp apply([%Function{value: function} | tail]), do: do_apply(function, tail)
  defp apply([function | tail]), do: do_apply(function, tail)

  defp do_apply(function, tail) do
    [{_type, ast, _meta} | reversed_args] = Enum.reverse(tail)
    args = Enum.reverse(reversed_args)
    func_args = Enum.concat(args, ast)
    function.(func_args)
  end

  defp symbol?([{:symbol, _}]), do: true
  defp symbol?(_), do: false

  defp vector?([{:vector, _ast, _meta}]), do: true
  defp vector?(_), do: false

  defp map?([{:map, _ast, _meta}]), do: true
  defp map?(_), do: false

  defp keyword([atom]) when is_atom(atom), do: atom
  defp keyword([atom]), do: String.to_atom(atom)

  defp cons([prepend, {_type, ast, meta}]), do: {:list, [prepend | ast], meta}

  defp concat(args) do
    args
      |> Enum.map(fn tuple -> elem(tuple, 1) end)
      |> Enum.concat
      |> list
  end

  defp assoc([{:map, hash_map, meta} | pairs]) do
    {:map, merge, _} = hash_map(pairs)
    {:map, Map.merge(hash_map, merge), meta}
  end

  defp dissoc([{:map, hash_map, meta} | keys]) do
    {:map, Map.drop(hash_map, keys), meta}
  end

  defp get([{:map, map, _}, key]), do: Map.get(map, key, nil)
  defp get(_), do: nil
end
