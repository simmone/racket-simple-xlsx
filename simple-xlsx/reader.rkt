
#lang racket

(provide (contract-out
          [with-input-from-xlsx-file (-> path-string? (-> any) any)]
          [sheet-name-rows (-> path-string? string? list?)]
          [get-sheet-names (-> list?)]
          [load-sheet (->* (string?) (procedure?) void?)]
          [load-sheet-ref (->* (natural?) (procedure?) void?)]
          [get-sheet-rows (-> list?)]
          ))

(require file/unzip)

(require "xlsx/xlsx.rkt")
(require "sheet/sheet.rkt")
(require "lib/dimension-lib.rkt")

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

(define (get-sheet-names)
  (let loop ([index 0]
             [result_list '()])
    (if (< index (XLSX-sheet_count (*CURRENT_XLSX*)))
        (loop
         (add1 index)
         (cons
          (hash-ref (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) index)
          result_list))
        (reverse result_list))))

(define (load-sheet sheet_name [user_procedure #f])
  (let ([sheet_index (hash-ref (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name 0)])
    (load-sheet-ref sheet_index user_procedure)))

(define (load-sheet-ref sheet_index [user_procedure #f])
  (let ([sheet 
         (if (regexp-match #rx"worksheets" (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) sheet_index))
             (load-data-sheet-file
              (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "xl" (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) sheet_index)))
             (load-chart-sheet-file
              (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "xl" (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) sheet_index))))])

  (set-XLSX-sheet_list!
   (*CURRENT_XLSX*)
   `(,@(XLSX-sheet_list (*CURRENT_XLSX*))
     ,sheet))

  (when user_procedure
        (parameterize
         ([*CURRENT_SHEET* sheet])
         (user_procedure)))))

(define (sheet-ref-row row_index)
  (let loop-col ([col_index 1]
                 [row '()])
    (if (<= col_index (cdr (DATA-SHEET-dimension (*CURRENT_SHEET*))))
        (loop-col
         (add1 col_index)
         (cons
          (sheet-ref-cell row_index col_index)
          row))
        (reverse row))))

(define (sheet-ref-cell row_index col_index)
  (sheet-cell (format "~a~a" (col_number->abc col_index) row_index) (*CURRENT_SHEET*)))

(define (sheet-cell item_name)
  (let ([sheet_map (get-field sheet_map (*CURRENT_XLSX*))]
        [data_type_map (get-field data_type_map (*CURRENT_XLSX*))]
        [shared_map (get-field shared_map (*CURRENT_XLSX*))])
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
         (DATA-SHEET-dimension sheet)])
    
    (let loop-row ([row_index 1]
                   [row_list '()])
      (if (<= row_index (car dimension))
          (loop-row
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
