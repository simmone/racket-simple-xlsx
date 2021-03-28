
#lang racket

(provide (contract-out
          [with-input-from-xlsx-file (-> path-string? (-> XLSX? any) any)]
          [sheet-name-rows (-> path-string? string? list?)]
          [get-sheet-names (-> XLSX? list?)]
          [load-sheet (->* (string? XLSX?) (procedure?) void?)]
          [load-sheet-ref (->* (natural? XLSX?) (procedure?) void?)]
          [get-sheet-rows (-> XLSX? list?)]
          ))

(require file/unzip)

(require "xlsx/xlsx.rkt")
(require "sheet/sheet.rkt")

(require "reader/load-workbook.rkt")
(require "reader/load-shared-strings.rkt")
(require "reader/load-workbook-rels.rkt")
(require "reader/load-sheet.rkt")

(define (with-input-from-xlsx-file xlsx_file user_proc)
  (call-with-unzip
   xlsx_file
   (lambda (xlsx_dir)
     (let ([_xlsx (new-xlsx)])
       (set-XLSX-xlsx_dir! _xlsx xlsx_dir)

       (load-workbook (build-path xlsx_dir "xl" "workbook.xml") _xlsx)

       (load-shared-strings (build-path xlsx_dir "xl" "sharedStrings.xml") _xlsx)

       (load-workbook-rels (build-path xlsx_dir "xl" "_rels" "workbook.xml.rels") _xlsx)

       (parameterize
        ([*CURRENT_XLSX* _xlsx])
        (user_proc))))))

(define (get-sheet-names xlsx)
  (let loop ([index 0]
             [result_list '()])
    (if (< index (XLSX-sheet_count xlsx))
        (loop
         (add1 index)
         (cons
          (hash-ref (XLSX-sheet_index_name_map xlsx) index)
          result_list))
        (reverse result_list))))

(define (load-sheet sheet_name xlsx [user_procedure #f])
  (let ([sheet_index (hash-ref (XLSX-sheet_name_index_map xlsx) sheet_name 0)])
    (load-sheet-ref sheet_index xlsx user_procedure)))

(define (load-sheet-ref sheet_index xlsx [user_procedure #f])
  (let ([sheet 
         (if (regexp-match #rx"worksheets" (hash-ref (XLSX-sheet_index_rel_map xlsx) sheet_index))
             (load-data-sheet-file
              (build-path (XLSX-xlsx_dir xlsx) "xl" (hash-ref (XLSX-sheet_index_rel_map xlsx) sheet_index)))
             (load-chart-sheet-file
              (build-path (XLSX-xlsx_dir xlsx) "xl" (hash-ref (XLSX-sheet_index_rel_map xlsx) sheet_index))))])

  (set-XLSX-sheet_list!
   xlsx
   `(,@(XLSX-sheet_list xlsx)
     ,sheet))

  (when user_procedure
        (parameterize
         ([*CURRENT_SHEET* sheet])
         (user_procedure)))))

(define (sheet-ref-row sheet row_index)
  (let loop-col ([col_index 1]
                 [row '()])
    (if (<= col_index (cdr (DATA-SHEET-dimension sheet)))
        (loop-col
         (add1 col_index)
         (cons
          (sheet-ref-cell row_index col_index sheet)
          row))
        (reverse row))))

(define (sheet-ref-cell row_index col_index sheet)
  (sheet-cell (format "~a~a" (number->abc col_index) row_index) sheet))

(define (sheet-cell item_name xlsx)
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


(define (get-rows sheet)
  (let ([dimension
         (XLSX-dimension sheet)])
    
    (let loop-row ([row_index 1]
                   [row_list '()])
      (if (<= row_index (car dimension))
          (loop 
           (add1 row_index)
           (cons
            (let loop-col ([col_index 1]
                           [col_list '()])
              (if (<= col_index (cdr dimension))
                  (loop-col
                   (add1 col_index)
                   (cons
                    (get-cell-value (string-append (number->abc col_index) (number->string row_index)) xlsx)
                    col_list))
                  (reverse col_list)))
            row_list))
          (reverse row_list)))))

(define (sheet-name-rows xlsx_file_path sheet_name)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet sheet_name xlsx)
     
     (get-sheet-rows xlsx))))
