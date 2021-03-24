#lang racket

(provide (contract-out
          [load-data-sheet-file (-> path-string? DATA-SHEET?)]
          [load-chart-sheet-file (-> path-string? CHART-SHEET?)]
          [get-sheet-rows (-> DATA-SHEET? list?)]
          ))

(require simple-xml)
(require file/unzip)

(require "../lib/lib.rkt")
(require "../sheet/sheet.rkt")
(require "../xlsx/range-lib.rkt")

(define (load-chart-sheet-file sheet_file)
  (void))

(define (load-data-sheet-file sheet_file)
  (let ([sheet 
         (DATA-SHEET
          '(0 . 0) (make-hash) (make-hash) (make-hash) (make-hash)
          (make-hash) (make-hash) '(0 . 0) (make-hash) (make-hash)
          (make-hash) (make-hash) (make-hash) (make-hash))])

    (let ([xml_hash (xml->hash sheet_file)])

      (set-DATA-SHEET-dimension!
       sheet
       (cons
        (hash-ref xml_hash "worksheet.sheetData.row's count")
        (hash-ref xml_hash "worksheet.cols.col's count")))

      (let loop-row ([row_count 1])
        (when (<= row_count (hash-ref xml_hash "worksheet.sheetData.row's count"))
          (let loop-col ([col_count 1])
            (when (<= col_count (hash-ref xml_hash "worksheet.cols.col's count"))
              (let (
                    [para_r (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.r" row_count col_count) "")]
                    [para_t (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.t" row_count col_count) #f)]
                    [para_s (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.s" row_count col_count) #f)]
                    [para_v (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.v" row_count col_count) #f)]
                    [para_f (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.f" row_count col_count) #f)]
                    )

                (hash-set! (DATA-SHEET-v_map sheet) para_r para_v)

                (when para_t (hash-set! (DATA-SHEET-t_map sheet) para_r para_t))

                (when para_f (hash-set! (DATA-SHEET-f_map sheet) para_r para_f))

                (when para_s (hash-set! (DATA-SHEET-s_map sheet) para_r para_s)))
              (loop-col (add1 col_count))))
          (loop-row (add1 row_count)))))
    sheet))

(define (get-cell-value item_name xlsx)
  (let ([sheet_map (get-field sheet_map xlsx)]
        [data_type_map (get-field data_type_map xlsx)]
        [shared_map (get-field shared_map xlsx)])
    (if (and
         (hash-has-key? sheet_map item_name)
         (not (null? (hash-ref data_type_map item_name))))
        (let* ([type (hash-ref data_type_map item_name)]
               [type_t (car type)]
               [type_s (cdr type)]
               [value (hash-ref sheet_map item_name)])
          (cond
           [(string=? type_t "s")
            (hash-ref shared_map value)]
           [(string=? type_t "n")
            (string->number value)]
           [(string=? type_t "")
            (string->number value)]))
        "")))

(define (get-cell-formula item_name xlsx)
  (let ([formula_map (get-field formula_map xlsx)])
    (if (hash-has-key? formula_map item_name)
        (hash-ref formula_map item_name)
        "")))

(define (get-sheet-dimension xlsx)
  (get-field dimension xlsx))

(define (get-sheet-rows xlsx)
  (let ([dimension null]
        [rows null]
        [cols null])
    
    (set! dimension (get-field dimension xlsx))
    
    (set! rows (car dimension))

    (set! cols (cdr dimension))

    (let loop ([row_index 1]
               [result_list '()])
      (if (<= row_index rows)
          (loop 
           (add1 row_index)
           (cons 
            (map
             (lambda (col_index)
               (get-cell-value (string-append (number->abc col_index) (number->string row_index)) xlsx))
             (number->list cols))
            result_list))
          (reverse result_list)))))
