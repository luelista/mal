\i init.sql
\i io.sql
\i types.sql
\i reader.sql
\i printer.sql
\i env.sql
\i core.sql

-- ---------------------------------------------------------
-- step1_read_print.sql

-- read
CREATE OR REPLACE FUNCTION READ(line varchar)
RETURNS integer AS $$
BEGIN
    RETURN read_str(line);
END; $$ LANGUAGE plpgsql;

-- eval
CREATE OR REPLACE FUNCTION is_pair(ast integer) RETURNS boolean AS $$
BEGIN
    RETURN _sequential_Q(ast) AND _count(ast) > 0;
END; $$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION quasiquote(ast integer) RETURNS integer AS $$
DECLARE
    a0   integer;
    a00  integer;
BEGIN
    IF NOT is_pair(ast) THEN
        RETURN _list(ARRAY[_symbolv('quote'), ast]);
    ELSE
        a0 := _nth(ast, 0);
        IF _symbol_Q(a0) AND a0 = _symbolv('unquote') THEN
            RETURN _nth(ast, 1);
        ELSE
            a00 := _nth(a0, 0);
            IF _symbol_Q(a00) AND a00 = _symbolv('splice-unquote') THEN
                RETURN _list(ARRAY[_symbolv('concat'),
                                   _nth(a0, 1),
                                   quasiquote(_rest(ast))]);
            END IF;
        END IF;
        RETURN _list(ARRAY[_symbolv('cons'),
                           quasiquote(_first(ast)),
                           quasiquote(_rest(ast))]);
    END IF;
END; $$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION is_macro_call(ast integer, env integer)
    RETURNS boolean AS $$
DECLARE
    a0      integer;
    f       integer;
    result  boolean = false;
BEGIN
    IF _list_Q(ast) THEN
        a0 = _first(ast);
        IF _symbol_Q(a0) AND env_find(env, _vstring(a0)) IS NOT NULL THEN
            f := env_get(env, a0);
            SELECT macro INTO result FROM collection
                WHERE collection_id = (SELECT collection_id FROM value
                                       WHERE value_id = f);
        END IF;
    END IF;
    RETURN result;
END; $$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION macroexpand(ast integer, env integer)
    RETURNS integer AS $$
DECLARE
    mac  integer;
BEGIN
    WHILE is_macro_call(ast, env)
    LOOP
        mac := env_get(env, _first(ast));
        ast := _apply(mac, _valueToArray(_rest(ast)));
    END LOOP;
    RETURN ast;
END; $$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION eval_ast(ast integer, env integer)
RETURNS integer AS $$
DECLARE
    type           integer;
    symkey         varchar;
    vid            integer;
    i              integer;
    src_coll_id    integer;
    dst_coll_id    integer = NULL;
    e              integer;
    result         integer;
BEGIN
    SELECT type_id INTO type FROM value WHERE value_id = ast;
    CASE
    WHEN type = 7 THEN
    BEGIN
        result := env_get(env, ast);
    END;
    WHEN type = 8 OR type = 9 THEN
    BEGIN
        src_coll_id := (SELECT collection_id FROM value WHERE value_id = ast);
        FOR vid, i IN (SELECT value_id, idx FROM collection
                       WHERE collection_id = src_coll_id)
        LOOP
            e := EVAL(vid, env);
            IF dst_coll_id IS NULL THEN
                dst_coll_id := COALESCE((SELECT Max(collection_id)
                                         FROM collection)+1,0);
            END IF;
            -- Evaluated each entry
            INSERT INTO collection (collection_id, idx, value_id)
                VALUES (dst_coll_id, i, e);
        END LOOP;
        -- Create value entry pointing to new collection
        INSERT INTO value (type_id, collection_id)
            VALUES (type, dst_coll_id)
            RETURNING value_id INTO result;
    END;
    ELSE
        result := ast;
    END CASE;

    RETURN result;
END; $$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION EVAL(ast integer, env integer)
RETURNS integer AS $$
DECLARE
    type     integer;
    a0       integer;
    a0sym    varchar;
    a1       integer;
    a2       integer;
    let_env  integer;
    binds    integer[];
    exprs    integer[];
    el       integer;
    fn       integer;
    fname    varchar;
    args     integer[];
    cond     integer;
    fast     integer;
    fparams  integer;
    fenv     integer;
    result   integer;
BEGIN
  LOOP
    --RAISE NOTICE 'EVAL: % [%]', pr_str(ast), ast;
    SELECT type_id INTO type FROM value WHERE value_id = ast;
    IF type <> 8 THEN
        RETURN eval_ast(ast, env);
    END IF;

    ast := macroexpand(ast, env);
    SELECT type_id INTO type FROM value WHERE value_id = ast;
    IF type <> 8 THEN
        RETURN eval_ast(ast, env);
    END IF;

    a0 := _first(ast);
    IF _symbol_Q(a0) THEN
        a0sym := (SELECT val_string FROM value WHERE value_id = a0);
    ELSE
        a0sym := '__<*fn*>__';
    END IF;

    --RAISE NOTICE 'ast: %, a0sym: %', ast, a0sym;
    CASE
    WHEN a0sym = 'def!' THEN
    BEGIN
        RETURN env_set(env, _nth(ast, 1), EVAL(_nth(ast, 2), env));
    END;
    WHEN a0sym = 'let*' THEN
    BEGIN
        let_env := env_new(env);
        a1 := _nth(ast, 1);
        binds := ARRAY(SELECT collection.value_id FROM collection INNER JOIN value
                       ON collection.collection_id=value.collection_id
                       WHERE value.value_id = a1
                       AND (collection.idx % 2) = 0
                       ORDER BY collection.idx);
        exprs := ARRAY(SELECT collection.value_id FROM collection INNER JOIN value
                       ON collection.collection_id=value.collection_id
                       WHERE value.value_id = a1
                       AND (collection.idx % 2) = 1
                       ORDER BY collection.idx);
        FOR idx IN array_lower(binds, 1) .. array_upper(binds, 1)
        LOOP
            PERFORM env_set(let_env, binds[idx], EVAL(exprs[idx], let_env));
        END LOOP;
        env := let_env;
        ast := _nth(ast, 2);
        CONTINUE; -- TCO
    END;
    WHEN a0sym = 'quote' THEN
    BEGIN
        RETURN _nth(ast, 1);
    END;
    WHEN a0sym = 'quasiquote' THEN
    BEGIN
        ast := quasiquote(_nth(ast, 1));
        CONTINUE; -- TCO
    END;
    WHEN a0sym = 'defmacro!' THEN
    BEGIN
        fn := EVAL(_nth(ast, 2), env);
        fn := _macro(fn);
        RETURN env_set(env, _nth(ast, 1), fn);
    END;
    WHEN a0sym = 'macroexpand' THEN
    BEGIN
        RETURN macroexpand(_nth(ast, 1), env);
    END;
    WHEN a0sym = 'try*' THEN
    BEGIN
        BEGIN
            RETURN EVAL(_nth(ast, 1), env);
            EXCEPTION WHEN OTHERS THEN
                IF _count(ast) >= 3 THEN
                    a2 = _nth(ast, 2);
                    IF _vstring(_nth(a2, 0)) = 'catch*' THEN
                        binds := ARRAY[_nth(a2, 1)];
                        exprs := ARRAY[_stringv(SQLERRM)];
                        env := env_new_bindings(env, _list(binds), exprs);
                        RETURN EVAL(_nth(a2, 2), env);
                    END IF;
                END IF;
                RAISE;
        END;
    END;
    WHEN a0sym = 'do' THEN
    BEGIN
        PERFORM eval_ast(_slice(ast, 1, _count(ast)-1), env);
        ast := _nth(ast, _count(ast)-1);
        CONTINUE; -- TCO
    END;
    WHEN a0sym = 'if' THEN
    BEGIN
        cond := EVAL(_nth(ast, 1), env);
        SELECT type_id INTO type FROM value WHERE value_id = cond;
        IF type = 0 OR type = 1 THEN -- nil or false
            IF _count(ast) > 3 THEN
                ast := _nth(ast, 3);
                CONTINUE; -- TCO
            ELSE
                RETURN 0; -- nil
            END IF;
        ELSE
            ast := _nth(ast, 2);
            CONTINUE; -- TCO
        END IF;
    END;
    WHEN a0sym = 'fn*' THEN
    BEGIN
        RETURN _function(_nth(ast, 2), _nth(ast, 1), env);
    END;
    ELSE
    BEGIN
        el := eval_ast(ast, env);
        SELECT type_id, collection_id, function_name
            INTO type, fn, fname
            FROM value WHERE value_id = _first(el);
        args := _restArray(el);
        IF type = 11 THEN
            EXECUTE format('SELECT %s($1);', fname)
                INTO result USING args;
            RETURN result;
        ELSIF type = 12 THEN
            SELECT value_id, params_id, env_id
                INTO fast, fparams, fenv
                FROM collection
                WHERE collection_id = fn;
            env := env_new_bindings(fenv, fparams, args);
            ast := fast;
            CONTINUE; -- TCO
        ELSE
            RAISE EXCEPTION 'Invalid function call';
        END IF;
    END;
    END CASE;
  END LOOP;
END; $$ LANGUAGE plpgsql;

-- print
CREATE OR REPLACE FUNCTION PRINT(exp integer) RETURNS varchar AS $$
BEGIN
    RETURN pr_str(exp);
END; $$ LANGUAGE plpgsql;


-- repl

-- repl_env is environment 0

CREATE OR REPLACE FUNCTION REP(line varchar)
RETURNS varchar AS $$
BEGIN
    RETURN PRINT(EVAL(READ(line), 0));
END; $$ LANGUAGE plpgsql;

-- core.sql: defined using SQL (in core.sql)
-- repl_env is created and populated with core functions in by core.sql
CREATE OR REPLACE FUNCTION mal_eval(args integer[]) RETURNS integer AS $$
BEGIN
    RETURN EVAL(args[1], 0);
END; $$ LANGUAGE plpgsql;
INSERT INTO value (type_id, function_name) VALUES (11, 'mal_eval');

SELECT env_vset(0, 'eval',
                   (SELECT value_id FROM value
                    WHERE function_name = 'mal_eval')) \g '/dev/null'
-- *ARGV* values are set by RUN
SELECT env_vset(0, '*ARGV*', READ('()'));


-- core.mal: defined using the language itself
SELECT REP('(def! not (fn* (a) (if a false true)))') \g '/dev/null'
SELECT REP('(def! load-file (fn* (f) (eval (read-string (str "(do " (slurp f) ")")))))') \g '/dev/null'
SELECT REP('(defmacro! cond (fn* (& xs) (if (> (count xs) 0) (list ''if (first xs) (if (> (count xs) 1) (nth xs 1) (throw "odd number of forms to cond")) (cons ''cond (rest (rest xs)))))))') \g '/dev/null'
SELECT REP('(defmacro! or (fn* (& xs) (if (empty? xs) nil (if (= 1 (count xs)) (first xs) `(let* (or_FIXME ~(first xs)) (if or_FIXME or_FIXME (or ~@(rest xs))))))))') \g '/dev/null'

CREATE OR REPLACE FUNCTION MAIN_LOOP()
RETURNS integer AS $$
DECLARE
    line    varchar;
    output  varchar;
BEGIN
    WHILE true
    LOOP
        BEGIN
            line := readline('user> ', 0);
            IF line IS NULL THEN RETURN 0; END IF;
            IF line <> '' THEN
                output := REP(line);
                PERFORM writeline(output);
            END IF;

            EXCEPTION WHEN OTHERS THEN
                PERFORM writeline('Error: ' || SQLERRM);
        END;
    END LOOP;
END; $$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION RUN(argstring varchar)
RETURNS void AS $$
DECLARE
    allargs  integer;
BEGIN
    allargs := READ(argstring);
    PERFORM env_vset(0, '*ARGV*', _rest(allargs));
    PERFORM REP('(load-file ' || pr_str(_first(allargs)) || ')');
    RETURN;
END; $$ LANGUAGE plpgsql;

