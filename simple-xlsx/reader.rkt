
#lang racket

(provide (contract-out
          [with-input-from-xlsx-file (-> path-string? (-> XLSX? any) any)]
          [sheet-name-rows (-> path-string? string? list?)]
          [get-sheet-names (-> XLSX? list?)]
          [load-sheet (-> string? XLSX? void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? XLSX? void?)]
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
       
       (user_proc _xlsx)))))

(define (load-sheet sheet_name xlsx)
  (let ([sheet_index (hash-ref (XLSX-sheet_name_index_map xlsx) sheet_name)])
    (load-sheet-ref sheet_index xlsx)))

(define (load-sheet-ref sheet_index xlsx)
  (set-XLSX-sheet_list!
   xlsx
   `(,@sheet_list
     ,(if (regexp-match #rx"worksheets" (XLSX-sheet_index_rel_map xlsx))
          (load-data-sheet-file
           (build-path (XLSX-xlsx_dir xlsx) "xl" (hash-ref (XLSX-sheet_index_rel_map xlsx) sheet_index)))
          (load-chart-sheet-file
           (build-path (XLSX-xlsx_dir xlsx) "xl" (hash-ref (XLSX-sheet_index_rel_map xlsx) sheet_index)))))))

(define (sheet-name-rows xlsx_file_path sheet_name)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet sheet_name xlsx)
     
     (get-sheet-rows xlsx))))
