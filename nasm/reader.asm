section .data

quote_symbol_string: db "quote"
.len: equ $ - quote_symbol_string
        
quasiquote_symbol_string: db "quasiquote"
.len: equ $ - quasiquote_symbol_string

unquote_symbol_string: db "unquote"
.len: equ $ - unquote_symbol_string

splice_unquote_symbol_string: db "splice-unquote"
.len: equ $ - splice_unquote_symbol_string
        
section .text

;; Read a string into memory as a form (nested lists and atoms)
;; Note: In this implementation the tokenizer is not done separately
;;
;; Input: Address of string (char array) in RSI
;;
;; Output: Address of object in RAX
;;
;; Uses registers:
;;  R12  Address of the start of the current list (starts 0)
;;  R13  Address of the current list tail
;;  R14  Stack pointer at start. Used for unwinding on error
;;  R15  Address of first list. Used for unwinding on error
;;
;; In addition, the tokenizer uses
;;
;;  RAX    (object return)
;;  RBX
;;  RCX    (character return in CL)
;;  RDX
;;  R8   ** State must be preserved
;;  R9   ** 
;;  R10  **
;;  R12
;;  R13
;;  R14   Original stack pointer on call
;;  R15   Top-level list, so all can be released on error
;;
read_str:
        ; Initialise tokenizer
        call tokenizer_init
        
        ; Set current list to zero
        mov r12, 0

        ; Set first list to zero
        mov r15, 0

        ; Save stack pointer for unwinding
        mov r14, rsp
        
.read_loop:

        call tokenizer_next
        cmp cl, 0
        jne .got_token

        ; Unexpected end of tokens
        jmp .unwind
        
.got_token:

        cmp cl, 'i'             ; An integer. Cons object in RAX
        je .finished
        cmp cl, '"'             ; A string. Array object in RAX
        je .finished
        cmp cl, 's'
        je .finished
        
        cmp cl, '('
        je .list_start
        
        cmp cl, ')'
        je .return_nil          ; Note: if reading a list, cl will be tested in the list reader

        cmp cl, 39              ; quote '
        je .handle_quote
        cmp cl, '`'
        je .handle_quasiquote
        cmp cl, '~'
        je .handle_unquote
        cmp cl, 1
        je .handle_splice_unquote
        
        ; Unknown
        jmp .return_nil
        
        ; --------------------------------
        
.list_start:
        
        ; Get the first value
        ; Note that we call rather than jmp because the first
        ; value needs to be treated differently. There's nothing
        ; to append to yet...
        call .read_loop
        
        ; rax now contains the first object
        cmp cl, ')'            ; Check if it was end of list
        jne .list_has_contents
        mov cl, 0               ; so ')' doesn't propagate to nested lists
        ret                    ; Returns 'nil' given "()"
.list_has_contents:
        ; If this is a Cons then use it
        ; If not, then need to allocate a Cons
        mov cl, BYTE [rax]
        mov ch, cl
        and ch, (block_mask + container_mask)   ; Tests block and container type
        jz .list_is_value

        ; If here then not a simple value, so need to allocate
        ; a Cons object
        
        ; Start new list
        push rax
        call alloc_cons         ; Address in rax
        pop rbx
        mov [rax], BYTE (block_cons + container_list + content_pointer)
        mov [rax + Cons.car], rbx
        ; Now have Cons in RAX, containing pointer to object as car
        
.list_is_value:
        ; Cons in RAX
        ; Make sure it's marked as a list
        mov cl, BYTE [rax]
        or cl, container_list
        mov [rax], BYTE cl

        mov r12, rax            ; Start of current list
        mov r13, rax            ; Set current list
        cmp r15, 0              ; Test if first list
        jne .list_read_loop
        mov r15, rax            ; Save the first, for unwinding
        
.list_read_loop:
        ; Repeatedly get the next value in the list
        ; (which may be other lists)
        ; until we get a ')' token

        push r12
        push r13
        call .read_loop         ; object in rax
        pop r13
        pop r12
        
        cmp cl, ')'            ; Check if it was end of list
        je .list_done          ; Have nil object in rax

        ; Test if this is a Cons value
        mov cl, BYTE [rax]
        mov ch, cl
        and ch, (block_mask + container_mask)   ; Tests block and container type
        jz .list_loop_is_value

        ; If here then not a simple value, so need to allocate
        ; a Cons object
        
        ; Start new list
        push rax
        call alloc_cons         ; Address in rax
        pop rbx
        mov [rax], BYTE (block_cons + container_list + content_pointer)
        mov [rax + Cons.car], rbx
        ; Now have Cons in RAX, containing pointer to object as car
        
.list_loop_is_value:
        ; Cons in RAX

        ; Make sure it's marked as a list
        mov cl, BYTE [rax]
        or cl, container_list
        mov [rax], BYTE cl
        
        ; Append to r13
        mov [r13 + Cons.typecdr], BYTE content_pointer
        mov [r13 + Cons.cdr], rax
        mov r13, rax            ; Set current list
        
        jmp .list_read_loop
        
.list_done:
        ; Release nil object in rax
        mov rsi, rax
        call release_cons
        
        ; Terminate the list
        mov [r13 + Cons.typecdr], BYTE content_nil
        mov QWORD [r13 + Cons.cdr], QWORD 0
        mov rax, r12            ; Start of current list
        
        ret

        ; --------------------------------
.handle_quote:
        ; Turn 'a into (quote a)
        call alloc_cons         ; Address in rax
        mov r12, rax

        ; Get a symbol "quote"
        push r8
        push r9
        mov rsi, quote_symbol_string
        mov edx, quote_symbol_string.len
        call raw_to_string      ; Address in rax
        pop r9
        pop r8

.wrap_next_object:
        mov [rax], BYTE maltype_symbol
        mov [r12], BYTE (block_cons + container_list + content_pointer)
        mov [r12 + Cons.car], rax
        
        ; Get the next object
        push r12
        call .read_loop         ; object in rax
        pop r12

        mov r13, rax            ; Put object to be quoted in r13

        call alloc_cons         ; Address in rax
        mov [rax], BYTE (block_cons + container_list + content_pointer)
        mov [rax + Cons.car], r13
        mov [rax + Cons.typecdr], BYTE content_nil
        
        ; Cons object in rax. Append to object in r12
        mov [r12 + Cons.typecdr], BYTE content_pointer
        mov [r12 + Cons.cdr], rax

        mov rax, r12
        ret

        ; --------------------------------
.handle_quasiquote:
        ; Turn `a into (quasiquote a)
        call alloc_cons         ; Address in rax
        mov r12, rax

        ; Get a symbol "quasiquote"
        push r8
        push r9
        mov rsi, quasiquote_symbol_string
        mov edx, quasiquote_symbol_string.len
        call raw_to_string      ; Address in rax
        pop r9
        pop r8
        jmp .wrap_next_object   ; From there the same as handle_quote

        ; --------------------------------
.handle_unquote:
        ; Turn ~a into (unquote a)
        call alloc_cons         ; Address in rax
        mov r12, rax

        ; Get a symbol "unquote"
        push r8
        push r9
        mov rsi, unquote_symbol_string
        mov edx, unquote_symbol_string.len
        call raw_to_string      ; Address in rax
        pop r9
        pop r8
        jmp .wrap_next_object   ; From there the same as handle_quote

        ; --------------------------------
.handle_splice_unquote:
        ; Turn ~@a into (unquote a)
        call alloc_cons         ; Address in rax
        mov r12, rax

        ; Get a symbol "unquote"
        push r8
        push r9
        mov rsi, splice_unquote_symbol_string
        mov edx, splice_unquote_symbol_string.len
        call raw_to_string      ; Address in rax
        pop r9
        pop r8
        jmp .wrap_next_object   ; From there the same as handle_quote
        
        ; --------------------------------
.finished:
        ret

.unwind:
        ; Jump to here cleans up

        mov rsp, r14            ; Rewind stack pointer
        cmp r15, 0              ; Check if there is a list
        je .return_nil
        mov rsi, r15
        call release_cons       ; releases everything recursively
        ; fall through to return_nil
.return_nil:
        ; Allocates a new Cons object with nil and returns
        ; Cleanup should happen before jumping here
        push rcx
        call alloc_cons
        pop rcx
        mov [rax], BYTE maltype_nil
        mov [rax + Cons.typecdr], BYTE content_nil
        ret

        
        
;; Initialise the tokenizer
;;
;; Input: Address of string in RSI
;; 
;; NOTE: This uses RSI, RAX and RBX, and expects these to be preserved
;; between calls to tokenizer_next_char
;;
;;  R9   Address of string
;;  R10  Position in data array
;;  R11  End of data array
;;
tokenizer_init:
        ; Save string to r9
        mov r9, rsi
        ; Put start of data array into r10
        mov r10, rsi
        add r10, Array.data
        ; Put end of data array into r11
        mov r11d, [rsi + Array.length] ; Length of array, zero-extended
        add r11, r10
        
        ret

;; Move onto the next chunk of the array
;; This is needed because strings are not stored in one
;; contiguous block of memory, but may use multiple Array
;; objects in a linked list
;;
;; If no chunks are left, then R10 = R11
tokenizer_next_chunk:
        mov r10, [r9 + Array.next]
        cmp r10, 0
        je .no_more
        ; More chunks left
        mov rsi, r10
        call tokenizer_init
        ret
.no_more:
        ; No more chunks left. R10 is zero
        mov r11, r10
        ret

;; Moves the next char into CL
;; If no more, puts 0 into CL
tokenizer_next_char:
        ; Check if we have reached the end of this chunk
        cmp r10, r11
        jne .chars_remain

        ; Hit the end. See if there is another chunk
        call tokenizer_next_chunk
        cmp r10, r11
        jne .chars_remain      ; Success, got another

        ; No more chunks
        mov cl, 0               ; Null char signals end
        ret
        
.chars_remain:
        mov cl, BYTE [r10]
        inc r10                 ; point to next byte
        ret
        
;; Get the next token
;; Token code is in CL register. Could be:
;; - 0 : Nil, finished
;; - Characters ()[]()'`~^@
;; - Pair '~@', represented by code 1
;; - A string: " in CL, and address in RAX
;; - An integer: 'i' in CL
;; - A symbol: 's' in CL, address in RAX
;;
;; Address of object in RAX
;;
;; May use registers:
;;    RBX
;;    RCX
;;    RDX
;;   
tokenizer_next:
        
.next_char:
        ; Fetch the next char into CL
        call tokenizer_next_char
        
        cmp cl, 0
        je .found               ; End, no more tokens
        
        ; Here expect to have:
        ; - The current character in CL
        ; - Address of next data in r10
        ; - Address of data end in r11

        ; Skip whitespace or commas
        cmp cl, ' '             ; Space
        je .next_char           
        cmp cl, ','             ; Comma
        je .next_char
        cmp cl, 9               ; Tab

        ; Special characters. These are returned in CL as-is
        cmp cl, '('
        je .found
        cmp cl, ')'
        je .found
        cmp cl, '['
        je .found
        cmp cl, ']'
        je .found
        cmp cl, '{'
        je .found
        cmp cl, '}'
        je .found
        cmp cl, 39              ; character '
        je .found
        cmp cl, 96              ; character `
        je .found
        cmp cl, '^'
        je .found
        cmp cl, '@'
        je .found
        cmp cl, '~'             ; Could be followed by '@'
        je .handle_tilde
        
        cmp cl, ';'             ; Start of a comment
        je .tokens_finished

        cmp cl, 34             ; Opening string quotes
        je .handle_string

        ; Could be number or symbol

        ; Check for a character 0-9
        cmp cl, '0'
        jl .handle_symbol
        cmp cl, '9'
        jg .handle_symbol

.handle_integer:
        ; Start integer
        ; accumulate in EDX
        xor edx, edx
        
.integer_loop:
        ; Here have a char 0-9 in CL
        sub cl, '0'            ; Convert to number between 0 and 9
        movzx ebx, cl
        add edx, ebx
        
        ; Push current state of the tokenizer
        push r9
        push r10
        push r11
        
        ; Peek at next character
        call tokenizer_next_char ; Next char in CL
        
        cmp cl, '0'
        jl .integer_finished
        cmp cl, '9'
        jg .integer_finished

        ; Discard old state by moving stack pointer
        add rsp, 24             ; 3 * 8 bytes
        
        imul edx, 10
        
        jmp .integer_loop
        
.integer_finished:
        ; Next char not an int
        
        ; Restore state of the tokenizer
        pop r11
        pop r10
        pop r9
        
        push rdx                ; Save the integer
        ; Get a Cons object to put the result into
        call alloc_cons
        
        pop rdx                 ; Restore integer
        
        ; Address of Cons now in RAX
        mov [rax], BYTE maltype_integer

        mov [rax + Cons.car], rdx
        
        mov cl, 'i'             ; Mark as an integer
        ret

        ; -------------------------------------------
.handle_symbol:
        ; Read characters until reaching whitespace, special character or end

        call string_new         
        mov rsi, rax  ; Output string in rsi
        
.symbol_loop:
        ; Put the current character into the array
        call string_append_char
        
        ; Push current state of the tokenizer
        push r9
        push r10
        push r11
        
        call tokenizer_next_char
        cmp cl, 0               ; End of characters
        je .symbol_finished
        
        cmp cl, ' '             ; Space
        je .symbol_finished           
        cmp cl, ','             ; Comma
        je .symbol_finished
        cmp cl, 9               ; Tab
        
        cmp cl, '('
        je .symbol_finished
        cmp cl, ')'
        je .symbol_finished
        cmp cl, '['
        je .symbol_finished
        cmp cl, ']'
        je .symbol_finished
        cmp cl, '{'
        je .symbol_finished
        cmp cl, '}'
        je .symbol_finished
        cmp cl, 39              ; character '
        je .symbol_finished
        cmp cl, 96              ; character `
        je .symbol_finished
        cmp cl, '^'
        je .symbol_finished
        cmp cl, '@'
        je .symbol_finished
        cmp cl, '~'
        je .symbol_finished
        cmp cl, ';'             ; Start of a comment
        je .symbol_finished
        cmp cl, 34              ; Opening string quotes
        je .symbol_finished

        ; Keeping current character
        ; Discard old state by moving stack pointer
        add rsp, 24             ; 3 * 8 bytes

        jmp .symbol_loop        ; Append to array
        
.symbol_finished:
        ; Not keeping current character
        ; Restore state of the tokenizer
        pop r11
        pop r10
        pop r9

        mov rax, rsi
        mov [rax], BYTE maltype_symbol ; Mark as a symbol
        mov cl, 's'                    ; used by read_str
        ret

        ; --------------------------------------------
.handle_string:
        ; Get an array to put the string into
        
        call string_new               ; Array in RAX
        
        ; Put start of data array into rbx
        mov rbx, rax
        add rbx, Array.data
        ; Put end of data array into rdx
        mov edx, DWORD [rax + Array.length] ; Length of array, zero-extended
        add rdx, rbx
        
        ; Now read chars from input string and push into output
.string_loop:
        
        call tokenizer_next_char
        cmp cl, 0               ; End of characters
        je .error
        
        cmp cl, 34              ; Finishing '"'
        je .string_done         ; Leave '"' in CL

        cmp cl, 92              ; Escape '\'
        jne .end_string_escape
        
        ; Current character is a '\'
        call tokenizer_next_char
        cmp cl, 0               ; End of characters
        je .error

        cmp cl, 'n'             ; \n, newline
        je .insert_newline

        ; Whatever is in cl is now put into string
        ; including '"'
        jmp .end_string_escape
        
.insert_newline:
        mov cl, 10
        jmp .end_string_escape
        
.end_string_escape:

        ; Put CL onto result array
        ; NOTE: this doesn't handle long strings (multiple memory blocks)
        mov [rbx], cl
        inc rbx

        jmp .string_loop

.string_done:
        ; Calculate the length from rbx
        sub rbx, Array.data
        sub rbx, rax
        mov [rax+Array.length], DWORD ebx
        ret
        
        ; ---------------------------------
        
.tokens_finished:
        mov cl, 0               ; End of tokens
        ret
        
.handle_tilde:
        ; Could have '~' or '~@'. Need to peek at the next char

        ; Push current state of the tokenizer
        push r9
        push r10
        push r11
        call tokenizer_next_char ; Next char in CL
        cmp cl, '@'
        jne .tilde_no_amp           ; Just '~', not '~@'
        ; Got '~@'
        mov cl, 1               ; Signals '~@'

        ; Discard old state by moving stack pointer
        add rsp, 24             ; 3 * 8 bytes
        ret
        
.tilde_no_amp:
        mov cl, '~'
        ; Restore state of the tokenizer
        pop r11
        pop r10
        pop r9
        ; fall through to found
        
.found:
        ret

.error:
        ret
        
