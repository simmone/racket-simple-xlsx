#lang racket

(require "../lib/lib.rkt")

(provide (contract-out
          [check-range (-> string? boolean?)]
          [check-col-range (-> string? string?)]
          [check-cell-range (-> string? boolean?)]
          [check-data-list (-> any/c boolean?)]
          [convert-range (-> string? string?)]
          [range-length (-> string? exact-nonnegative-integer?)]
          ))

(define (check-data-list data_list)
  (when (not (list? data_list))
        (error "data is not list type"))
  
  (when (equal? data_list '())
        (error "data has no children list"))
  
  (let loop ([loop_list data_list]
             [child_length -1])
    (when (not (null? loop_list))
          (when (not (list? (car loop_list)))
                (error "data's children is not list type"))
          
          (when (and
                 (not (= child_length -1))
                 (not (= child_length (length (car loop_list)))))
                (error "data's children's length is not consistent."))
          
          (loop (cdr loop_list) (length (car loop_list)))))
  #t)

(define (check-range range_str)
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
           [start_col_index (second items)]
           [end_col_index (third items)])
      (if (string>? start_col_index end_col_index)
          (error (format "col index should from small to big[~a]" col_range_str))
          (string-append start_col_index "-" end_col_index)))]
   [else
    (error (format "invalid col range! should be like this: A-Z or 1-10 or A or 1 but is [~a]" col_range_str))]
   ))

(define (check-cell-range cell_range_str)
  (if (regexp-match #rx"^[A-Z]+[0-9]+-[A-Z]+[0-9]+$" cell_range_str)
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
          #t]))
      (error (format "invalid cell range! should be like this: A1-B2[~a]" cell_range_str))))

(define (convert-range range_str)
  (when (check-range range_str)
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
