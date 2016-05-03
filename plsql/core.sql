PROMPT 'core.sql start';

CREATE OR REPLACE TYPE core_ns_type IS TABLE OF varchar2(100);
/

CREATE OR REPLACE PACKAGE core IS

FUNCTION do_core_func(M IN OUT NOCOPY mem_type,
                      H IN OUT NOCOPY types.map_entry_table,
                      fn integer,
                      a mal_seq_items_type) RETURN integer;

FUNCTION get_core_ns RETURN core_ns_type;

END core;
/


CREATE OR REPLACE PACKAGE BODY core AS

-- general functions
FUNCTION equal_Q(M IN OUT NOCOPY mem_type,
                 H IN OUT NOCOPY types.map_entry_table,
                 args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.tf(types.equal_Q(M, H, args(1), args(2)));
END;

-- scalar functiosn
FUNCTION symbol(M IN OUT NOCOPY mem_type,
                val integer) RETURN integer IS
BEGIN
    RETURN types.symbol(M, TREAT(M(val) AS mal_str_type).val_str);
END;

FUNCTION keyword(M IN OUT NOCOPY mem_type,
                 val integer) RETURN integer IS
BEGIN
    IF types.string_Q(M, val) THEN
        RETURN types.keyword(M, TREAT(M(val) AS mal_str_type).val_str);
    ELSIF types.keyword_Q(M, val) THEN
        RETURN val;
    ELSE
        raise_application_error(-20009,
            'invalid keyword call', TRUE);
    END IF;
END;


-- string functions
FUNCTION pr_str(M IN OUT NOCOPY mem_type,
                H IN OUT NOCOPY types.map_entry_table,
                args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.string(M, printer.pr_str_seq(M, H, args, ' ', TRUE));
END;

FUNCTION str(M IN OUT NOCOPY mem_type,
             H IN OUT NOCOPY types.map_entry_table,
             args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.string(M, printer.pr_str_seq(M, H, args, '', FALSE));
END;

FUNCTION prn(M IN OUT NOCOPY mem_type,
             H IN OUT NOCOPY types.map_entry_table,
             args mal_seq_items_type) RETURN integer IS
BEGIN
    stream_writeline(printer.pr_str_seq(M, H, args, ' ', TRUE));
    RETURN 1;  -- nil
END;

FUNCTION println(M IN OUT NOCOPY mem_type,
                 H IN OUT NOCOPY types.map_entry_table,
                 args mal_seq_items_type) RETURN integer IS
BEGIN
    stream_writeline(printer.pr_str_seq(M, H, args, ' ', FALSE));
    RETURN 1;  -- nil
END;

FUNCTION read_string(M IN OUT NOCOPY mem_type,
                     H IN OUT NOCOPY types.map_entry_table,
                     args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN reader.read_str(M, H, TREAT(M(args(1)) AS mal_str_type).val_str);
END;

FUNCTION readline(M IN OUT NOCOPY mem_type,
                  prompt integer) RETURN integer IS
    input  varchar2(4000);
BEGIN
    input := stream_readline(TREAT(M(prompt) AS mal_str_type).val_str, 0);
    RETURN types.string(M, input);
EXCEPTION WHEN OTHERS THEN
    IF SQLCODE = -20001 THEN  -- io streams closed
        RETURN 1;  -- nil
    ELSE
        RAISE;
    END IF;
END;

FUNCTION slurp(M IN OUT NOCOPY mem_type,
               args mal_seq_items_type) RETURN integer IS
    content  varchar2(4000);
BEGIN
    -- stream_writeline('here1: ' || TREAT(args(1) AS mal_str_type).val_str);
    content := file_open_and_read(TREAT(M(args(1)) AS mal_str_type).val_str);
    content := REPLACE(content, '\n', chr(10));
    RETURN types.string(M, content);
END;


-- numeric functions
FUNCTION lt(M IN OUT NOCOPY mem_type,
            args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.tf(TREAT(M(args(1)) AS mal_int_type).val_int <
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION lte(M IN OUT NOCOPY mem_type,
             args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.tf(TREAT(M(args(1)) AS mal_int_type).val_int <=
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION gt(M IN OUT NOCOPY mem_type,
            args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.tf(TREAT(M(args(1)) AS mal_int_type).val_int >
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION gte(M IN OUT NOCOPY mem_type,
             args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.tf(TREAT(M(args(1)) AS mal_int_type).val_int >=
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION add(M IN OUT NOCOPY mem_type,
             args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.int(M, TREAT(M(args(1)) AS mal_int_type).val_int +
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION subtract(M IN OUT NOCOPY mem_type,
                  args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.int(M, TREAT(M(args(1)) AS mal_int_type).val_int -
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION multiply(M IN OUT NOCOPY mem_type,
                  args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.int(M, TREAT(M(args(1)) AS mal_int_type).val_int *
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION divide(M IN OUT NOCOPY mem_type,
                args mal_seq_items_type) RETURN integer IS
BEGIN
    RETURN types.int(M, TREAT(M(args(1)) AS mal_int_type).val_int /
                        TREAT(M(args(2)) AS mal_int_type).val_int);
END;

FUNCTION time_ms(M IN OUT NOCOPY mem_type) RETURN integer IS
    now  integer;
BEGIN
    -- SELECT (SYSDATE - TO_DATE('01-01-1970 00:00:00', 'DD-MM-YYYY HH24:MI:SS')) * 24 * 60 * 60 * 1000
    --    INTO now FROM DUAL;
    SELECT extract(day from(sys_extract_utc(systimestamp) - to_timestamp('1970-01-01', 'YYYY-MM-DD'))) * 86400000 + to_number(to_char(sys_extract_utc(systimestamp), 'SSSSSFF3'))
        INTO now
        FROM dual;
    RETURN types.int(M, now);
END;

-- hash-map functions
FUNCTION assoc(M IN OUT NOCOPY mem_type,
               H IN OUT NOCOPY types.map_entry_table,
               hm integer,
               kvs mal_seq_items_type) RETURN integer IS
    new_hm    integer;
    midx      integer;
BEGIN
    new_hm := types.clone(M, H, hm);
    midx := TREAT(M(new_hm) AS mal_map_type).map_idx;
    -- Add the new key/values
    midx := types.assoc_BANG(M, H, midx, kvs);
    RETURN new_hm;
END;

FUNCTION dissoc(M IN OUT NOCOPY mem_type,
                H IN OUT NOCOPY types.map_entry_table,
                hm integer,
                ks mal_seq_items_type) RETURN integer IS
    new_hm    integer;
    midx      integer;
BEGIN
    new_hm := types.clone(M, H, hm);
    midx := TREAT(M(new_hm) AS mal_map_type).map_idx;
    -- Remove the keys
    midx := types.dissoc_BANG(M, H, midx, ks);
    RETURN new_hm;
END;


FUNCTION get(M IN OUT NOCOPY mem_type,
             H IN OUT NOCOPY types.map_entry_table,
             hm integer, key integer) RETURN integer IS
    midx  integer;
    k     varchar2(256);
    val   integer;
BEGIN
    IF M(hm).type_id = 0 THEN
        RETURN 1;  -- nil
    END IF;
    midx := TREAT(M(hm) AS mal_map_type).map_idx;
    k := TREAT(M(key) AS mal_str_type).val_str;
    IF H(midx).EXISTS(k) THEN
        RETURN H(midx)(k);
    ELSE
        RETURN 1;  -- nil
    END IF;
END;

FUNCTION contains_Q(M IN OUT NOCOPY mem_type,
             H IN OUT NOCOPY types.map_entry_table,
             hm integer, key integer) RETURN integer IS
    midx  integer;
    k     varchar2(256);
    val   integer;
BEGIN
    midx := TREAT(M(hm) AS mal_map_type).map_idx;
    k := TREAT(M(key) AS mal_str_type).val_str;
    RETURN types.tf(H(midx).EXISTS(k));
END;

FUNCTION keys(M IN OUT NOCOPY mem_type,
              H IN OUT NOCOPY types.map_entry_table,
              hm integer) RETURN integer IS
    midx  integer;
    k     varchar2(256);
    ks    mal_seq_items_type;
    val   integer;
BEGIN
    midx := TREAT(M(hm) AS mal_map_type).map_idx;
    ks := mal_seq_items_type();

    k := H(midx).FIRST();
    WHILE k IS NOT NULL LOOP
        ks.EXTEND();
        ks(ks.COUNT()) := types.string(M, k);
        k := H(midx).NEXT(k);
    END LOOP;

    RETURN types.seq(M, 8, ks);
END;

FUNCTION vals(M IN OUT NOCOPY mem_type,
              H IN OUT NOCOPY types.map_entry_table,
              hm integer) RETURN integer IS
    midx  integer;
    k     varchar2(256);
    ks    mal_seq_items_type;
    val   integer;
BEGIN
    midx := TREAT(M(hm) AS mal_map_type).map_idx;
    ks := mal_seq_items_type();

    k := H(midx).FIRST();
    WHILE k IS NOT NULL LOOP
        ks.EXTEND();
        ks(ks.COUNT()) := H(midx)(k);
        k := H(midx).NEXT(k);
    END LOOP;

    RETURN types.seq(M, 8, ks);
END;


-- sequence functions
FUNCTION cons(M IN OUT NOCOPY mem_type,
              args mal_seq_items_type) RETURN integer IS
    new_items  mal_seq_items_type;
    len        integer;
    i          integer;
BEGIN
    new_items := mal_seq_items_type();
    len := types.count(M, args(2));
    new_items.EXTEND(len+1);
    new_items(1) := args(1);
    FOR i IN 1..len LOOP
        new_items(i+1) := TREAT(M(args(2)) AS mal_seq_type).val_seq(i);
    END LOOP;
    RETURN types.seq(M, 8, new_items);
END;

FUNCTION concat(M IN OUT NOCOPY mem_type,
                args mal_seq_items_type) RETURN integer IS
    new_items  mal_seq_items_type;
    cur_len    integer;
    seq_len    integer;
    i          integer;
    j          integer;
BEGIN
    new_items := mal_seq_items_type();
    cur_len := 0;
    FOR i IN 1..args.COUNT() LOOP
        seq_len := types.count(M, args(i));
        new_items.EXTEND(seq_len);
        FOR j IN 1..seq_len LOOP
            new_items(cur_len + j) := types.nth(M, args(i), j-1);
        END LOOP;
        cur_len := cur_len + seq_len;
    END LOOP;
    RETURN types.seq(M, 8, new_items);
END;


FUNCTION nth(M IN OUT NOCOPY mem_type,
             val integer,
             ival integer) RETURN integer IS
    idx  integer;
BEGIN
    idx := TREAT(M(ival) AS mal_int_type).val_int;
    RETURN types.nth(M, val, idx);
END;

FUNCTION first(M IN OUT NOCOPY mem_type,
               val integer) RETURN integer IS
BEGIN
    IF val = 1 OR types.count(M, val) = 0 THEN
        RETURN 1;  -- nil
    ELSE
        RETURN types.first(M, val);
    END IF;
END;

FUNCTION rest(M IN OUT NOCOPY mem_type,
              val integer) RETURN integer IS
BEGIN
    IF val = 1 OR types.count(M, val) = 0 THEN
        RETURN types.list(M);
    ELSE
        RETURN types.slice(M, val, 1);
    END IF;
END;

FUNCTION do_count(M IN OUT NOCOPY mem_type,
               val integer) RETURN integer IS
BEGIN
    IF M(val).type_id = 0 THEN
        RETURN types.int(M, 0);
    ELSE
        RETURN types.int(M, types.count(M, val));
    END IF;
END;


FUNCTION conj(M IN OUT NOCOPY mem_type,
              seq integer,
              vals mal_seq_items_type) RETURN integer IS
    type_id  integer;
    slen     integer;
    items    mal_seq_items_type;
BEGIN
    type_id := M(seq).type_id;
    slen := types.count(M, seq);
    items := mal_seq_items_type();
    items.EXTEND(slen + vals.COUNT());
    CASE
    WHEN type_id = 8 THEN
        FOR i IN 1..vals.COUNT() LOOP
            items(i) := vals(vals.COUNT + 1 - i);
        END LOOP;
        FOR i IN 1..slen LOOP
            items(vals.COUNT() + i) := types.nth(M, seq, i-1);
        END LOOP;
    WHEN type_id = 9 THEN
        FOR i IN 1..slen LOOP
            items(i) := types.nth(M, seq, i-1);
        END LOOP;
        FOR i IN 1..vals.COUNT() LOOP
            items(slen + i) := vals(i);
        END LOOP;
    ELSE
        raise_application_error(-20009,
            'conj: not supported on type ' || type_id, TRUE);
    END CASE;
    RETURN types.seq(M, type_id, items);
END;

FUNCTION seq(M IN OUT NOCOPY mem_type,
             val integer) RETURN integer IS
    type_id    integer;
    new_val    integer;
    str        varchar2(4000);
    str_items  mal_seq_items_type;
BEGIN
    type_id := M(val).type_id;
    CASE
    WHEN type_id = 8 THEN
        IF types.count(M, val) = 0 THEN
            RETURN 1;  -- nil
        END IF;
        RETURN val;
    WHEN type_id = 9 THEN
        IF types.count(M, val) = 0 THEN
            RETURN 1;  -- nil
        END IF;
        RETURN types.seq(M, 8, TREAT(M(val) AS mal_seq_type).val_seq);
    WHEN types.string_Q(M, val) THEN
        str := TREAT(M(val) AS mal_str_type).val_str;
        IF str IS NULL THEN
            RETURN 1;  -- nil
        END IF;
        str_items := mal_seq_items_type();
        str_items.EXTEND(LENGTH(str));
        FOR i IN 1..LENGTH(str) LOOP
            str_items(i) := types.string(M, SUBSTR(str, i, 1));
        END LOOP;
        RETURN types.seq(M, 8, str_items);
    WHEN type_id = 0 THEN
        RETURN 1;  -- nil
    ELSE
        raise_application_error(-20009,
            'seq: not supported on type ' || type_id, TRUE);
    END CASE;
END;

-- atom functions
FUNCTION reset_BANG(M IN OUT NOCOPY mem_type,
                    atm integer,
                    new_val integer) RETURN integer IS
BEGIN
    M(atm) := mal_atom_type(13, new_val);
    RETURN new_val;
END;

-- metadata functions
FUNCTION meta(M IN OUT NOCOPY mem_type,
              val integer) RETURN integer IS
    type_id  integer;
BEGIN
    type_id := M(val).type_id;
    IF type_id IN (8,9) THEN  -- list/vector
        RETURN TREAT(M(val) AS mal_seq_type).meta;
    ELSIF type_id = 10 THEN   -- hash-map
        RETURN TREAT(M(val) AS mal_map_type).meta;
    ELSIF type_id = 11 THEN   -- native function
        RETURN 1;  -- nil
    ELSIF type_id = 12 THEN   -- mal function
        RETURN TREAT(M(val) AS malfunc_type).meta;
    ELSE
        raise_application_error(-20006,
            'meta: metadata not supported on type', TRUE);
    END IF;
END;

-- general native function case/switch
FUNCTION do_core_func(M IN OUT NOCOPY mem_type,
                      H IN OUT NOCOPY types.map_entry_table,
                      fn integer,
                      a mal_seq_items_type) RETURN integer IS
    fname  varchar(100);
    idx    integer;
BEGIN
    IF M(fn).type_id <> 11 THEN
        raise_application_error(-20004,
            'Invalid function call', TRUE);
    END IF;

    fname := TREAT(M(fn) AS mal_str_type).val_str;

    CASE
    WHEN fname = '='           THEN RETURN equal_Q(M, H, a);

    WHEN fname = 'nil?'        THEN RETURN types.tf(a(1) = 1);
    WHEN fname = 'false?'      THEN RETURN types.tf(a(1) = 2);
    WHEN fname = 'true?'       THEN RETURN types.tf(a(1) = 3);
    WHEN fname = 'string?'     THEN RETURN types.tf(types.string_Q(M, a(1)));
    WHEN fname = 'symbol'      THEN RETURN symbol(M, a(1));
    WHEN fname = 'symbol?'     THEN RETURN types.tf(M(a(1)).type_id = 7);
    WHEN fname = 'keyword'     THEN RETURN keyword(M, a(1));
    WHEN fname = 'keyword?'    THEN RETURN types.tf(types.keyword_Q(M, a(1)));

    WHEN fname = 'pr-str'      THEN RETURN pr_str(M, H, a);
    WHEN fname = 'str'         THEN RETURN str(M, H, a);
    WHEN fname = 'prn'         THEN RETURN prn(M, H, a);
    WHEN fname = 'println'     THEN RETURN println(M, H, a);
    WHEN fname = 'read-string' THEN RETURN read_string(M, H, a);
    WHEN fname = 'readline'    THEN RETURN readline(M, a(1));
    WHEN fname = 'slurp'       THEN RETURN slurp(M, a);

    WHEN fname = '<'           THEN RETURN lt(M, a);
    WHEN fname = '<='          THEN RETURN lte(M, a);
    WHEN fname = '>'           THEN RETURN gt(M, a);
    WHEN fname = '>='          THEN RETURN gte(M, a);
    WHEN fname = '+'           THEN RETURN add(M, a);
    WHEN fname = '-'           THEN RETURN subtract(M, a);
    WHEN fname = '*'           THEN RETURN multiply(M, a);
    WHEN fname = '/'           THEN RETURN divide(M, a);
    WHEN fname = 'time-ms'     THEN RETURN time_ms(M);

    WHEN fname = 'list'        THEN RETURN types.seq(M, 8, a);
    WHEN fname = 'list?'       THEN RETURN types.tf(M(a(1)).type_id = 8);
    WHEN fname = 'vector'      THEN RETURN types.seq(M, 9, a);
    WHEN fname = 'vector?'     THEN RETURN types.tf(M(a(1)).type_id = 9);
    WHEN fname = 'hash-map'    THEN RETURN types.hash_map(M, H, a);
    WHEN fname = 'assoc'       THEN RETURN assoc(M, H, a(1), types.islice(a, 1));
    WHEN fname = 'dissoc'      THEN RETURN dissoc(M, H, a(1), types.islice(a, 1));
    WHEN fname = 'map?'        THEN RETURN types.tf(M(a(1)).type_id = 10);
    WHEN fname = 'get'         THEN RETURN get(M, H, a(1), a(2));
    WHEN fname = 'contains?'   THEN RETURN contains_Q(M, H, a(1), a(2));
    WHEN fname = 'keys'        THEN RETURN keys(M, H, a(1));
    WHEN fname = 'vals'        THEN RETURN vals(M, H, a(1));

    WHEN fname = 'sequential?' THEN RETURN types.tf(M(a(1)).type_id IN (8,9));
    WHEN fname = 'cons'        THEN RETURN cons(M, a);
    WHEN fname = 'concat'      THEN RETURN concat(M, a);
    WHEN fname = 'nth'         THEN RETURN nth(M, a(1), a(2));
    WHEN fname = 'first'       THEN RETURN first(M, a(1));
    WHEN fname = 'rest'        THEN RETURN rest(M, a(1));
    WHEN fname = 'empty?'      THEN RETURN types.tf(0 = types.count(M, a(1)));
    WHEN fname = 'count'       THEN RETURN do_count(M, a(1));

    WHEN fname = 'conj'        THEN RETURN conj(M, a(1), types.islice(a, 1));
    WHEN fname = 'seq'         THEN RETURN seq(M, a(1));

    WHEN fname = 'meta'        THEN RETURN meta(M, a(1));
    WHEN fname = 'with-meta'   THEN RETURN types.clone(M, H, a(1), a(2));
    WHEN fname = 'atom'        THEN RETURN types.atom_new(M, a(1));
    WHEN fname = 'atom?'       THEN RETURN types.tf(M(a(1)).type_id = 13);
    WHEN fname = 'deref'       THEN RETURN TREAT(M(a(1)) AS mal_atom_type).val;
    WHEN fname = 'reset!'      THEN RETURN reset_BANG(M, a(1), a(2));

    ELSE raise_application_error(-20004, 'Invalid function call', TRUE);
    END CASE;
END;

FUNCTION get_core_ns RETURN core_ns_type IS
BEGIN
    RETURN core_ns_type(
        '=',
        'throw',

        'nil?',
        'true?',
        'false?',
        'string?',
        'symbol',
        'symbol?',
        'keyword',
        'keyword?',

        'pr-str',
        'str',
        'prn',
        'println',
        'read-string',
        'readline',
        'slurp',

        '<',
        '<=',
        '>',
        '>=',
        '+',
        '-',
        '*',
        '/',
        'time-ms',

        'list',
        'list?',
        'vector',
        'vector?',
        'hash-map',
        'assoc',
        'dissoc',
        'map?',
        'get',
        'contains?',
        'keys',
        'vals',

        'sequential?',
        'cons',
        'concat',
        'nth',
        'first',
        'rest',
        'empty?',
        'count',
        'apply',   -- defined in step do_builtin function
        'map',     -- defined in step do_builtin function

        'conj',
        'seq',

        'meta',
        'with-meta',
        'atom',
        'atom?',
        'deref',
        'reset!',
        'swap!'    -- defined in step do_builtin function
    );
END;

END core;
/
show errors;

PROMPT 'core.sql finished';
