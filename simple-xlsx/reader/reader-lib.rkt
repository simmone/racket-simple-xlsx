#lang racket

(provide (contract-out
          [get-range-data-ref (-> natural? string? list?)]
          ))

(require "../xlsx/xlsx.rkt")

(define (get-range-data-ref sheet_index range_str)
  (when (check-data-range-valid #:sheet_name sheet_name #:range_str range_str)
    (let* ([data_sheet (get-sheet-by-name sheet_name)]
           [col_range (regexp-match* #rx"([A-Z]+)" range_str)]
           [start_col_name (first col_range)]
           [col_start_index (sub1 (abc->number start_col_name))]
           [end_col_name (second col_range)]
           [col_end_index (sub1 (abc->number end_col_name))]
           [row_range (regexp-match* #rx"([0-9]+)" range_str)]
           [row_start_index (string->number (first row_range))]
           [row_end_index (string->number (second row_range))]
           [direction (if (string=? start_col_name end_col_name) 'vertical 'horizontal)])
      (if (eq? direction 'vertical)
          (let loop ([loop_list (DATA-SHEET-rows data_sheet)]
                     [row_count 1]
                     [result_list '()])
            (if (not (null? loop_list))
                (if (and (>= row_count row_start_index) (<= row_count row_end_index))
                    (loop (cdr loop_list) (add1 row_count) 
                          (cons (list-ref (car loop_list) col_start_index) result_list))
                    (loop (cdr loop_list) (add1 row_count) result_list))
                (reverse result_list)))
          (let loop ([loop_list (list-ref (DATA-SHEET-rows data_sheet) (sub1 row_start_index))]
                     [col_index 0]
                     [result_list '()])
            (if (not (null? loop_list))
                (if (and (>= col_index col_start_index) (<= col_index col_end_index))
                    (loop (cdr loop_list) (add1 col_index) (cons (car loop_list) result_list))
                    (loop (cdr loop_list) (add1 col_index) result_list))
                (reverse result_list)))))))
