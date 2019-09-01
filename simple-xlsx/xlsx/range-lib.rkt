#lang racket

(require "../lib/lib.rkt")

(provide (contract-out
          [abc->number (-> string? natural?)]
          [abc->range (-> string? pair?)]
          [number->abc (-> natural? string?)]
          [get-dimension (-> list? string?)]
          [check-cell-range (-> string? (or/c #f string?))]
          [check-col-range (-> string? string?)]
          [check-row-range (-> string? string?)]
          [only-one-row/col-data? (-> string? boolean?)]
          [convert-range (-> string? string?)]
          [range-length (-> string? natural?)]
          [range-to-cell-hash (-> string? any/c hash?)]
          [range-to-row-hash (-> string? any/c hash?)]
          [range-to-col-hash (-> string? any/c hash?)]
          [combine-cols-hash (-> hash? hash? list?)]
          [combine-hash-in-hash (-> (listof hash?) hash?)]
          [cell->rowcol (-> string? (cons/c natural? natural?))]
          [cross-cell-style (-> hash? hash? symbol? hash?)]
          [expand-row-style-to-cell (-> hash? hash? void?)]
          [expand-col-style-to-cell (-> hash? hash? void?)]
          ))

(define (abc->number abc)
  (let ([sum 0])
    (let loop ([char_list (reverse (string->list abc))]
               [base 0])
      (when (not (null? char_list))
            (let* ([alpha (car char_list)]
                   [alpha_int (add1 (- (char->integer alpha) (char->integer #\A)))])
              (set! sum (+ (* alpha_int (expt 26 base)) sum)))
            (loop (cdr char_list) (add1 base))))
    sum))

(define (abc->range abc_range)
  (cond
   [(regexp-match #rx"^([0-9]+|[A-Z]+)-([0-9]+|[A-Z]+)$" abc_range)
    (let ([abc_items (regexp-split #rx"-" abc_range)])
      (if (= (length abc_items) 2)
          (let* ([first_item (first abc_items)]
                 [second_item (second abc_items)]
                 [start_index
                  (cond
                   [(regexp-match #rx"^[0-9]+$" first_item)
                    (string->number first_item)]
                   [(regexp-match #rx"^[A-Z]+$" first_item)
                    (abc->number first_item)]
                   [else
                    1])]
                 [end_index
                  (cond
                   [(regexp-match #rx"^[0-9]+$" second_item)
                    (string->number second_item)]
                   [(regexp-match #rx"^[A-Z]+$" second_item)
                    (abc->number second_item)]
                   [else
                    1])])
            (if (<= start_index end_index)
                (cons start_index end_index)
                (cons 1 1)))
          (cons 1 1)))]
   [(regexp-match #rx"^[0-9]+$" abc_range)
    (cons (string->number abc_range) (string->number abc_range))]
   [(regexp-match #rx"^[A-Z]+$" abc_range)
    (cons (abc->number abc_range) (abc->number abc_range))]
   [else
    (cons 1 1)]))

(define (number->abc num)
  (let ([abc ""])
    (let loop ([loop_num num])
      (if (> loop_num 26)
          (let-values ([(quo remain) (quotient/remainder loop_num 26)])
            (if (= remain 0)
                (begin
                  (set! abc (string-append "Z" abc))
                  (loop (sub1 quo)))
                (begin
                  (set! abc (string-append (string (integer->char (+ 64 remain))) abc))
                  (loop quo))))
          (set! abc (string-append (string (integer->char (+ 64 loop_num))) abc))))
    abc))

(define (get-dimension data_list)
  (let ([rows (length data_list)]
        [cols 0])
    (let loop ([loop_list data_list])
      (when (not (null? loop_list))
            (when (> (length (car loop_list)) cols)
                  (set! cols (length (car loop_list))))
            (loop (cdr loop_list))))

    (string-append (number->abc cols) (number->string rows))))

(define (check-cell-range cell_range_str)
  (cond
   [(regexp-match #rx"^[A-Z]+[0-9]+-[A-Z]+[0-9]+$" cell_range_str)
    (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" cell_range_str)]
           [start_col_name (second items)]
           [start_col_index (string->number (third items))]
           [end_col_name (fourth items)]
           [end_col_index (string->number (fifth items))])
      (cond
       [(string<? end_col_name start_col_name)
        (error (format "col name should from small to big[~a]" cell_range_str))]
       [(< end_col_index start_col_index)
        (error (format "col index should from small to big[~a]" cell_range_str))]
       [else
        cell_range_str]))]
   [(regexp-match #rx"^[A-Z]+[0-9]+$" cell_range_str)
    (format "~a-~a" cell_range_str cell_range_str)]
   [else
    (error (format "invalid cell range! should be like this: A1-B2[~a]" cell_range_str))]))

(define (check-col-range col_range_str)
  (cond
   [(regexp-match #rx"^[A-Z]+$" col_range_str)
    (let* ([items (regexp-match #rx"^([A-Z]+)$" col_range_str)]
           [start_col_name (second items)]
           [end_col_name start_col_name])
      (string-append start_col_name "-" end_col_name))]
   [(regexp-match #rx"^[0-9]+$" col_range_str)
    (let* ([items (regexp-match #rx"^([0-9]+)$" col_range_str)]
           [start_col_index (second items)]
           [end_col_index start_col_index])
      (string-append start_col_index "-" end_col_index))]
   [(regexp-match #rx"^[A-Z]+-[A-Z]+$" col_range_str)
    (let* ([items (regexp-match #rx"^([A-Z]+)-([A-Z]+)$" col_range_str)]
           [start_col_name (second items)]
           [end_col_name (third items)])
      (if (string>? start_col_name end_col_name)
          (error (format "col name should from small to big[~a]" col_range_str))
          (string-append start_col_name "-" end_col_name)))]
   [(regexp-match #rx"^[0-9]+-[0-9]+$" col_range_str)
    (let* ([items (regexp-match #rx"^([0-9]+)-([0-9]+)$" col_range_str)]
           [start_col_index (string->number (second items))]
           [end_col_index (string->number (third items))])
      (if (> start_col_index end_col_index)
          (error (format "col index should from small to big[~a]" col_range_str))
          (format "~a-~a" start_col_index end_col_index)))]
   [else
    (error (format "invalid col range! should be like this: A-Z or 1-10 or A or 1 but is [~a]" col_range_str))]
   ))

(define (check-row-range row_range_str)
  (cond
   [(regexp-match #rx"^[0-9]+$" row_range_str)
    (let* ([items (regexp-match #rx"^([0-9]+)$" row_range_str)]
           [start_row_index (second items)]
           [end_row_index start_row_index])
      (string-append start_row_index "-" end_row_index))]
   [(regexp-match #rx"^[0-9]+-[0-9]+$" row_range_str)
    (let* ([items (regexp-match #rx"^([0-9]+)-([0-9]+)$" row_range_str)]
           [start_row_index (string->number (second items))]
           [end_row_index (string->number (third items))])
      (if (> start_row_index end_row_index)
          (error (format "row index should from small to big[~a]" row_range_str))
          (format "~a-~a" start_row_index end_row_index)))]
   [else
    (error (format "invalid row range! should be like this: A-Z or 1-10 or A or 1 but is [~a]" row_range_str))]
   ))

(define (only-one-row/col-data? range_str)
  (if (regexp-match #rx"^[A-Z]+[0-9]+-[A-Z]+[0-9]+$" range_str)
      (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
             [start_col_name (second items)]
             [start_row_index (third items)]
             [end_col_name (fourth items)]
             [end_row_index (fifth items)])

        (if (string=? start_col_name end_col_name)
            (if (> (string->number start_row_index) (string->number end_row_index))
                (error (format "range's direction is vertical, index is invalid.[~a][~a]" start_row_index end_row_index))
                #t)
            (if (string=? start_row_index end_row_index)
                (if (> (abc->number start_col_name) (abc->number end_col_name))
                    (error (format "range's direction is horizontal, col name is invalid.[~a][~a]" start_col_name end_col_name))
                    #t)
                (error (format "range's direction confused. should like A1-A20 or A2-Z2, but get ~a" range_str)))))
      (error (format "range format should like A1-A20 or A2-Z2, but get ~a" range_str))))

(define (convert-range range_str)
  (when (check-cell-range range_str)
        (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
               [start_col_name (second items)]
               [start_index (third items)]
               [end_col_name (fourth items)]
               [end_index (fifth items)])
          (string-append "$" start_col_name "$" start_index ":$" end_col_name "$" end_index))))

(define (range-length range_str)
  (let* ([items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
         [start_col_name (second items)]
         [start_row_index (third items)]
         [end_col_name (fourth items)]
         [end_row_index (fifth items)])

    (if (string=? start_col_name end_col_name)
        (add1 (- (string->number end_row_index) (string->number start_row_index)))
        (add1 (- (abc->number end_col_name) (abc->number start_col_name))))))

(define (range-to-cell-hash range_str val)
  (let ([flat_map (make-hash)])
    (when (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)
          (let* ([range_items (regexp-match #rx"^([A-Z]+)([0-9]+)-([A-Z]+)([0-9]+)$" range_str)]
                 [start_col_index (abc->number (list-ref range_items 1))]
                 [start_row_index (string->number (list-ref range_items 2))]
                 [end_col_index (abc->number (list-ref range_items 3))]
                 [end_row_index (string->number (list-ref range_items 4))])
            (let range-loop ([loop_col_index start_col_index]
                             [loop_row_index start_row_index])
              (when (and
                     (<= loop_col_index end_col_index)
                     (<= loop_row_index end_row_index))
                    (hash-set! flat_map 
                               (string-append (number->abc loop_col_index) (number->string loop_row_index))
                               val)
                    (cond
                     [(< loop_col_index end_col_index)
                      (range-loop (add1 loop_col_index) loop_row_index)]
                     [(< loop_row_index end_row_index)
                      (range-loop start_col_index (add1 loop_row_index))])))))
    flat_map))

(define (range-to-row-hash range_str val)
  (let ([flat_map (make-hash)])
    (let* ([items (regexp-split #rx"-" range_str)]
           [start_row (string->number (first items))]
           [end_row (string->number (second items))])
      (when (and
             (natural? start_row)
             (natural? end_row)
             (>= end_row start_row))
            (let loop ([loop_row start_row])
              (when (<= loop_row end_row)
                    (hash-set! flat_map loop_row val)
                    (loop (add1 loop_row)))))
      flat_map)))

(define (range-to-col-hash range_str val)
  (let ([flat_map (make-hash)])
    (let* ([items (abc->range range_str)]
           [start_col (car items)]
           [end_col (cdr items)])
      (when (and
             (natural? start_col)
             (natural? end_col)
             (>= end_col start_col))
            (let loop ([loop_col start_col])
              (when (<= loop_col end_col)
                    (hash-set! flat_map loop_col val)
                    (loop (add1 loop_col)))))
      flat_map)))

(define (combine-cols-hash width_map style_map)
  (let ([col_hash (make-hash)]
        [width_list (sort (hash->list width_map) < #:key car)]
        [style_list (sort (hash->list style_map) < #:key car)])
    
    (let loop-width ([loop_list width_list])
      (when (not (null? loop_list))
            (hash-set! col_hash (caar loop_list) (list (cdar loop_list) #f))
            (loop-width (cdr loop_list))))
    
    (let loop-style ([loop_list style_list])
      (when (not (null? loop_list))
            (let ([val_list (hash-ref col_hash (caar loop_list) (list #f #f))])
              (hash-set! col_hash (caar loop_list) (list-set val_list 1 (cdar loop_list))))
            (loop-style (cdr loop_list))))

    (let ([sorted_range_val_list (sort (hash->list col_hash) < #:key car)])
      (let loop-squeeze ([loop_list sorted_range_val_list]
                         [last_val (cdar sorted_range_val_list)]
                         [last_range_start (caar sorted_range_val_list)]
                         [last_range_end (caar sorted_range_val_list)]
                         [next_step (caar sorted_range_val_list)]
                         [squeeze_list '()])
        (if (not (null? loop_list))
            (cond
             [(and
               (equal? (cdar loop_list) last_val)
               (= (caar loop_list) next_step))
              (loop-squeeze
               (cdr loop_list)
               last_val
               last_range_start
               next_step
               (add1 (caar loop_list))
               squeeze_list)]
             [(and
               (equal? (cdar loop_list) last_val)
               (not (= (caar loop_list) next_step)))
              (loop-squeeze
               (cdr loop_list)
               last_val
               (caar loop_list)
               (caar loop_list)
               (add1 (caar loop_list))
               (cons (cons (cons last_range_start last_range_end) last_val) squeeze_list))]
             [(or
               (and
                (not (equal? (cdar loop_list) last_val))
                (= (caar loop_list) next_step))
               (and
                (not (equal? (cdar loop_list) last_val))
                (not (= (caar loop_list) next_step))))
              (loop-squeeze
               (cdr loop_list)
               (cdar loop_list)
               (caar loop_list)
               (caar loop_list)
               (add1 (caar loop_list))
               (cons (cons (cons last_range_start last_range_end) last_val) squeeze_list))])
            (reverse (cons (cons (cons last_range_start last_range_end) last_val) squeeze_list)))))))

(define (combine-hash-in-hash hash_list)
  (let ([result_map (make-hash)])
    (let outer-hash-loop ([hashes hash_list])
      (when (not (null? hashes))
            (hash-for-each
             (car hashes)
             (lambda (cell_name style_hash)
               (if (hash-has-key? result_map cell_name)
                   (let ([old_hash (hash-copy (hash-ref result_map cell_name))])

                     (hash-for-each
                      style_hash
                      (lambda (ik iv)
                        (hash-set! old_hash ik iv)))
                     (hash-set! result_map cell_name old_hash))
                   (hash-set! result_map cell_name style_hash))))
            (outer-hash-loop (cdr hashes))))
    result_map))

(define (cell->rowcol cell_str)
  (let ([split_items (regexp-match #rx"^([a-zA-Z]+)([0-9]+)$" cell_str)])
    (if split_items
        (cons (string->number (third split_items)) (abc->number (second split_items)))
        '(0 . 0))))

(define (cross-cell-style row_map col_map type)
  (let ([cross_cell_map (make-hash)])
    (hash-for-each
     row_map
     (lambda (row_seq row_style_map)
       (hash-for-each
        col_map
        (lambda (col_seq col_style_map)
          (let ([head_style_map (if (eq? type 'row) col_style_map row_style_map)]
                [tail_style_map (if (eq? type 'row) row_style_map col_style_map)]
                [style_map (make-hash)])
            (hash-for-each
             head_style_map
             (lambda (style val)
               (hash-set! style_map style val)))

            (hash-for-each
             tail_style_map
             (lambda (style val)
               (hash-set! style_map style val)))
            
            (hash-set! cross_cell_map (string-append (number->abc col_seq) (number->string row_seq)) style_map))))))
    cross_cell_map))

(define (expand-row-style-to-cell row_style_map cell_style_map)
  (hash-for-each
   cell_style_map
   (lambda (cell cell_style_map)
     (let ([rowcol (cell->rowcol cell)])
       (when (hash-has-key? row_style_map (car rowcol))
         (hash-for-each
          (hash-ref row_style_map (car rowcol))
          (lambda (style val)
            (hash-set! cell_style_map style val))))))))

(define (expand-col-style-to-cell col_style_map cell_style_map)
  (hash-for-each
   cell_style_map
   (lambda (cell cell_style_map)
     (let ([rowcol (cell->rowcol cell)])
       (when (hash-has-key? col_style_map (cdr rowcol))
         (hash-for-each
          (hash-ref col_style_map (cdr rowcol))
          (lambda (style val)
            (hash-set! cell_style_map style val))))))))
