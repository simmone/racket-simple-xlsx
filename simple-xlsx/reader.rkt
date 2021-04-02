
#lang racket

(provide (contract-out
          [with-input-from-xlsx-file (-> path-string? (-> any) any)]
          [xlsx-sheet-count (-> natural?)]
          [xlsx-sheet-names (-> list?)]
          [load-sheet (->* (string?) (procedure?) void?)]
          [load-sheet-ref (->* (natural?) (procedure?) void?)]
          [sheet-dimension (-> (cons/c natural? natural?))]
          [get-rows (-> list?)]
          [get-cell-value (-> string? any)]
          [sheet-name-rows (-> path-string? string? list?)]
          [sheet-ref-rows (-> path-string? natural? list?)]
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

(define (xlsx-sheet-count)
  (XLSX-sheet_count (*CURRENT_XLSX*)))

(define (xlsx-sheet-names)
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

(define (get-cell-value item_name)
  (let ([r (string-upcase item_name)]
        [rvtsf_map (DATA-SHEET-rvtsf_map (*CURRENT_SHEET*))]
        [shared_strings_map (XLSX-shared_strings_map (*CURRENT_XLSX*))])
    (if (hash-has-key? rvtsf_map r)
        (let* ([vtsf (hash-ref rvtsf_map r)]
               [v (first vtsf)]
               [t (second vtsf)]
               [s (third vtsf)])
          (cond
           [(string=? t "s")
            (hash-ref shared_strings_map (string->number v))]
           [(or (string=? t "n") (string=? t ""))
            (string->number v)]
           [else
            ""]))
        "")))

(define (sheet-dimension)
  (DATA-SHEET-dimension (*CURRENT_SHEET*)))

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

(define (get-rows)
  (let ([dimension
         (DATA-SHEET-dimension (*CURRENT_SHEET*))])
    
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
                    (get-cell-value (string-append (col_number->abc col_index) (number->string row_index)))
                    col_list))
                  (reverse col_list)))
            row_list))
          (reverse row_list)))))

(define (sheet-name-rows xlsx_file_path sheet_name)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet 
      sheet_name
      (lambda ()
        (get-rows))))))

(define (sheet-ref-rows xlsx_file_path sheet_index)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet-ref
      sheet_index
      (lambda ()
        (get-rows))))))

