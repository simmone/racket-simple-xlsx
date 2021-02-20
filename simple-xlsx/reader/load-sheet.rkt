#lang racket

(provide (contract-out
          [load-sheet (-> string? (is-a?/c read-xlsx%) void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? (is-a?/c read-xlsx%) void?)]
          [get-sheet-rows (-> (is-a?/c read-xlsx%) list?)]
          [sheet-name-rows (-> path-string? string? list?)]
          [sheet-ref-rows (-> path-string? exact-nonnegative-integer? list?)]
          ))

(require simple-xml)
(require file/unzip)

(require "lib/lib.rkt")
(require "xlsx/xlsx.rkt")
(require "xlsx/range-lib.rkt")

(define (load-sheet sheet_name xlsx)
  (let ([data_map (make-hash)]
        [formula_map (make-hash)]
        [type_map (make-hash)]
        [dimension_col 0]
        [dimension ""]
        [data_sheet_file_name
         (build-path (get-field xlsx_dir xlsx) "xl" (hash-ref (get-field relation_name_map xlsx) (hash-ref (get-field sheet_name_map xlsx) sheet_name)))])

    (let ([xml_hash (xml->hash sheet_xml_file)])
      (set-field! dimension xlsx (cons
                                  (hash-ref xml_hash "worksheet.sheetData.row's count")
                                  (hash-ref xml_hash "worksheet.cols's count")))

      (let loop-row ([row_count 1])
        (when (<= row_count (hash-ref xml_hash "worksheet.sheetData.row's count"))
          (let loop-col ([col_count 1])
            (when (<= col_count (hash-ref xml_hash "worksheet.sheetData.cols's count"))
              (let (
                    [para_r (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.r" row_count col_count) "")]
                    [para_t (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.t" row_count col_count) "")]
                    [para_s (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.s" row_count col_count) "")]
                    [para_v (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.v" row_count col_count) #f)]
                    [para_f (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.f" row_count col_count) #f)]
                    )

                (hash-set! type_map para_r (cons para_t para_s))

                (when para_v
                  (hash-set! data_map para_r para_v))

                (when para_f
                  (hash-set! formula_map para_r para_f))
                )
              (loop-col (add1 col_count))))
          (loop-row (add1 row_count))))

                                      (let loop-cell ([cell_list (cdr cell_item)])
                                        (when (not (null? cell_list))
                                              (cond 
                                                [(equal? (caar cell_list) 'v)
                                                 (hash-set! data_map para_r (caddar cell_list))]
                                                [(equal? (caar cell_list) 'f)
                                                 (hash-set! formula_map para_r (caddar cell_list))]
                                              )
                                              (loop-cell (cdr cell_list))))
                                      )))))
                    row_xml)))
           rows)
          )

    (set-field! sheet_map xlsx data_map)
    (set-field! formula_map xlsx formula_map)
    (set-field! data_type_map xlsx type_map)))

(define (load-sheet-ref sheet_index xlsx)
  (load-sheet (list-ref (get-sheet-names xlsx) sheet_index) xlsx))

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

(define (sheet-name-rows xlsx_file_path sheet_name)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet sheet_name xlsx)
     
     (get-sheet-rows xlsx))))

(define (sheet-ref-rows xlsx_file_path sheet_index)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet-ref sheet_index xlsx)
     
     (get-sheet-rows xlsx))))
