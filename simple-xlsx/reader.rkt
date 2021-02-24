#lang racket

(provide (contract-out
          [with-input-from-xlsx-file (-> path-string? (-> (is-a?/c read-xlsx%) any) any)]
          [get-sheet-names (-> (is-a?/c read-xlsx%) list?)]
          [get-cell-value (-> string? (is-a?/c read-xlsx%) any)]
          [get-cell-formula (-> string? (is-a?/c read-xlsx%) string?)]
          [get-sheet-dimension (-> (is-a?/c read-xlsx%) pair?)]
          [load-sheet (-> string? (is-a?/c read-xlsx%) void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? (is-a?/c read-xlsx%) void?)]
          [get-sheet-rows (-> (is-a?/c read-xlsx%) list?)]
          [sheet-name-rows (-> path-string? string? list?)]
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
     (let (
           [read_sheet_id_list '()]
           [read_sheet_id_name_map (make-hash)]
           [read_sheet_name_id_map (make-hash)]
           [read_sheet_id_rid_map (make-hash)]
           [read_shared_strings_map #f]
           [read_sheet_rid_rel_map #f]
           [xlsx_obj #f])

       (let-values ([(_sheet_id_list _sheet_id_rid_map _sheet_name_id_map _sheet_id_name_map)
                     (load-workbook (build-path tmp_dir "xl" "workbook.xml"))])
         (set! read_sheet_id_list _sheet_id_list)
         (set! read_sheet_id_rid_map _sheet_id_rid_map)
         (set! read_sheet_name_id_map _sheet_name_id_map)
         (set! read_sheet_id_name_map _sheet_id_name_map))

       (set! read_shared_strings_map (load-shared-strings (build-path tmp_dir "xl" "sharedStrings.xml")))

       (set! read_sheet_rid_rel_map (load-workbook-rels (build-path tmp_dir "xl" "_rels" "workbook.xml.rels")))
       
       (set! xlsx_obj
             (new new-xlsx%
                  (xlsx_dir tmp_dir)
                  (sheet_id_list read_sheet_id_list)
                  (sheet_id_name_map read_sheet_id_name_map)
                  (sheet_name_id_map read_sheet_name_id_map)
                  (sheet_id_rid_map read_sheet_id_rid_map)
                  (sheet_rid_rel_map read_sheet_rid_rel_map)
                  (sheet #f)
                  (style #f)
                  (shared_strings_map read_shared_strings_map)
                  (formula_map #f)
                  (data_type_map #f)
                  (dimension #f)))

       (user_proc xlsx_obj)))))

(define (sheet-name-rows xlsx_file_path sheet_name)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet sheet_name xlsx)
     
     (get-sheet-rows xlsx))))



