#!/bin/bash

RL_HISTORY_FILE=${HOME}/.mal-history
SKIP_INIT="${SKIP_INIT:-}"

SQLPATH=$(readlink -f $(dirname $0))
SQLPLUS="sqlplus -S system/oracle"

# Load the SQL code
if [ -z "${SKIP_INIT}" ]; then
    out=$(echo "" | ${SQLPLUS} @$1)
    if echo "${out}" | grep -vs "^No errors.$" \
        | grep -si error >/dev/null; then
    #if echo "${out}"  | grep -si error >/dev/null; then
        echo "${out}"
        exit 1
    fi
fi

# open I/O streams
echo -e "BEGIN stream_open(0); stream_open(1); END;\n/" \
    | ${SQLPLUS} >/dev/null

# Stream from table to stdout
(
while true; do
    out="$(echo "SELECT stream_read(1) FROM dual;" \
        | ${SQLPLUS} 2>/dev/null)" || break
    #echo "out: [${out}]"
    echo "${out}"
done
) &

# Perform readline input into stream table when requested
(
[ -r ${RL_HISTORY_FILE} ] && history -r ${RL_HISTORY_FILE}
while true; do
    prompt=$(echo "SELECT stream_wait_rl_prompt(0) FROM dual;" \
        | ${SQLPLUS} 2>/dev/null) || break
    # Prompt is returned single-quoted because sqlplus trims trailing
    # whitespace. Remove the single quotes from the beginning and end:
    prompt=${prompt%\'}
    prompt=${prompt#\'}
    #echo "prompt: [${prompt}]"

    IFS= read -u 0 -r -e -p "${prompt}" line || break
    if [ "${line}" ]; then
        history -s -- "${line}"        # add to history
        history -a ${RL_HISTORY_FILE}  # save history to file
    fi

    # Escape (double) single quotes per SQL norm
    line=${line//\'/\'\'}
    #echo "line: [${line}]"
    ( echo -n "BEGIN stream_writeline('${line}', 0); END;";
      echo -en "\n/" ) \
        | ${SQLPLUS} >/dev/null || break
done
echo -e "BEGIN stream_close(0); stream_close(1); END;\n/" \
    | ${SQLPLUS} > /dev/null
) <&0 >&1 &

res=0
shift
if [ $# -gt 0 ]; then
    # If there are command line arguments then run a command and exit
    args=$(for a in "$@"; do echo -n "\"$a\" "; done)
    echo -e "BEGIN MAIN('$(pwd)', :'args'); END;\n/" \
        | ${SQLPLUS} "(${args})" > /dev/null
    res=$?
else
    # Start main loop in the background
    echo "SELECT mal_pkg.MAIN('$(pwd)') FROM dual;" \
        | ${SQLPLUS} > /dev/null
    res=$?
fi
echo -e "BEGIN stream_close(0); stream_close(1); END;\n/" \
    | ${SQLPLUS} > /dev/null
exit ${res}
