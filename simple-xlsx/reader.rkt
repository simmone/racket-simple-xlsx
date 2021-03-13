#lang racket

(provide (contract-out
          [with-input-from-xlsx-file (-> path-string? (-> (is-a?/c read-xlsx%) any) any)]
          [sheet-name-rows (-> path-string? string? list?)]
          [get-sheet-names (-> (is-a?/c read-xlsx%) list?)]
          [get-cell-value (-> string? (is-a?/c read-xlsx%) any)]
          [get-cell-formula (-> string? (is-a?/c read-xlsx%) string?)]
          [get-sheet-dimension (-> (is-a?/c read-xlsx%) pair?)]
          [load-sheet (-> string? (is-a?/c read-xlsx%) void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? (is-a?/c read-xlsx%) void?)]
          [get-sheet-rows (-> (is-a?/c read-xlsx%) list?)]
          [sheet-ref-rows (-> path-string? exact-nonnegative-integer? list?)]
          ))

(require "xlsx/xlsx.rkt")

(require "reader/load-workbook.rkt")
(require "reader/load-shared-string.rkt")
(require "reader/load-workbook-rels.rkt")
(require "reader/load-sheet.rkt")

(define (with-input-from-xlsx-file xlsx_file user_proc)
  (call-with-unzip
   xlsx_file
   (lambda (tmp_dir)
     (let ([_xlsx (new-xlsx tmp_dir)])
       (load-workbook (build-path tmp_dir "xl" "workbook.xml") _xlsx)

       (load-shared-strings (build-path tmp_dir "xl" "sharedStrings.xml") _xlsx)

       (load-workbook-rels (build-path tmp_dir "xl" "_rels" "workbook.xml.rels") _xlsx)
       
       (user_proc _xlsx)))))

(define (sheet-name-rows xlsx_file_path sheet_name)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet sheet_name xlsx)
     
     (get-sheet-rows xlsx))))
