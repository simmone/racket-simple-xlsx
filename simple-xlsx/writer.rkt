#lang racket

(provide (contract-out
          [xlsx-data% class?]
          [xlsx-data? any/c]
          [col-attr any/c]
          [write-xlsx-file (-> xlsx-data? path-string? void?)]
          ))

;; data list:
;; '(((1 2 3 4) (1 2 3) (1 3 4)) ((3 4 5 6) (3 6 7 8)))
;; first level children is sheets
;; each sheet contains rows, row's length is not same
;; sheet name list:
;; '("Sheet1" "Sheet2" ...)
;; #f use default Sheet1, Sheet2... as sheet name

(require racket/date)

(require "lib/lib.rkt")
(require "writer/content-type.rkt")
(require "writer/_rels/rels.rkt")
(require "writer/docProps/docprops-app.rkt")
(require "writer/docProps/docprops-core.rkt")
(require "writer/xl/_rels/workbook-xml-rels.rkt")
(require "writer/xl/printerSettings/printerSettings.rkt")
(require "writer/xl/sharedStrings.rkt")
(require "writer/xl/styles.rkt")
(require "writer/xl/theme/theme.rkt")
(require "writer/xl/workbook.rkt")
(require "writer/xl/worksheets/_rels/rels.rkt")
(require "writer/xl/worksheets/worksheet.rkt")

;; represent a column's attibutes, if not want set a specific attr, set it to #f
;; example: (col-attr #f "red") means only set color
(struct col-attr (width color))

;; col_attr_hash used to set col's attribute: '#hash(("A" . (col-attr 100 "white")))
(define xlsx-data%
  (class object%
         (field (sheet_data_list '()))
         (field (sheet_name_list '()))
         (field (sheet_attr_hash (make-hash)))

         (define/public (add-sheet sheet_row_list sheet_name #:col_attr_hash [col_attr_hash #f])
           (set! sheet_data_list `(,@sheet_data_list ,sheet_row_list))
           (set! sheet_name_list `(,@sheet_name_list ,sheet_name))
           (when col_attr_hash
                 (hash-set! sheet_attr_hash (length sheet_data_list) col_attr_hash)))

         (super-new)))

(define (xlsx-data? obj)
  (and (field-bound? sheet_data_list obj) (field-bound? sheet_name_list obj)))

(define (write-xlsx-file xlsx_data xlsx_file_name)
  (when (file-exists? xlsx_file_name)
        (delete-file xlsx_file_name))
  
  (let ([tmp_dir #f])
    (dynamic-wind
        (lambda () (set! tmp_dir (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
        (lambda ()
          (let ([sheet_data_list (get-field sheet_data_list xlsx_data)]
                [sheet_name_list (get-field sheet_name_list xlsx_data)])
            (let-values ([(string_index_list string_index_map) (get-string-index sheet_data_list)])
              (let ([sheet_count (length sheet_data_list)])
                ;; [Content_Types].xml
                (write-content-type-file tmp_dir sheet_count)

                ;; _rels
                (let ([rels_dir (build-path tmp_dir "_rels")])
                  (make-directory* rels_dir)
                  (write-rels-file rels_dir))

                ;; docProps
                (let ([doc_props_dir (build-path tmp_dir "docProps")])
                  (make-directory* doc_props_dir)
                  (write-docprops-app-file doc_props_dir sheet_name_list)
                  (write-docprops-core-file doc_props_dir (current-date)))
                
                ;; xl
                (let ([xl_dir (build-path tmp_dir "xl")])
                  ;; _rels
                  (let ([rels_dir (build-path xl_dir "_rels")])
                    (make-directory* rels_dir)
                    (write-workbook-xml-rels-file rels_dir sheet_count))
                  
                  ;; printerSettings
                  (let ([printer_settings_dir (build-path xl_dir "printerSettings")])
                    (make-directory* printer_settings_dir)
                    (create-printer-settings printer_settings_dir sheet_count))

                  ;; sharedStrings
                  (write-shared-strings-file xl_dir string_index_list)

                  ;; styles.xml
                  (write-styles-file xl_dir)

                  ;; theme
                  (let ([theme_dir (build-path xl_dir "theme")])
                    (make-directory* theme_dir)
                    (write-theme-file theme_dir))

                  ;; workbook
                  (write-workbook-file xl_dir sheet_name_list)

                  ;; worksheets
                  (let ([worksheets_dir (build-path xl_dir "worksheets")])
                    ;; _rels
                    (let ([worksheets_rels_dir (build-path worksheets_dir "_rels")])
                      (make-directory* worksheets_rels_dir)
                      (write-worksheets-rels-file worksheets_rels_dir sheet_count))

                    ;; worksheet
                    (let loop ([sheets_data sheet_data_list]
                               [index 1])
                      (when (not (null? sheets_data))
                            (write-sheet-file worksheets_dir index (car sheets_data) string_index_map)
                            (loop (cdr sheets_data) (add1 index))))
                    )
                  )
                ))
            (zip-xlsx xlsx_file_name tmp_dir)))
          (lambda ()
            (delete-directory/files tmp_dir)))))
