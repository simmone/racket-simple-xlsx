#lang racket

(provide (contract-out 
          [write-xlsx-file (-> list? (or/c list? #f) path-string? void?)]
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

(define (write-xlsx-file data_list sheet_name_list file_name)
  (let ([tmp_dir #f])
    (dynamic-wind
        (lambda () (set! tmp_dir (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
        (lambda ()
          (let-values ([(string_index_list string_index_map) (get-string-index data_list)])
            (let* ([sheet_count (length data_list)]
                   [real_sheet_name_list (if sheet_name_list sheet_name_list (create-sheet-name-list sheet_count))])
              ;; [Content_Types].xml
              (write-content-type-file tmp_dir sheet_count)

              ;; _rels
              (let ([rels_dir (build-path tmp_dir "_rels")])
                (make-directory* rels_dir)
                (write-rels-file rels_dir))

              ;; docProps
              (let ([doc_props_dir (build-path tmp_dir "docProps")])
                (make-directory* doc_props_dir)
                (write-docprops-app-file doc_props_dir real_sheet_name_list)
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
                (write-workbook-file xl_dir real_sheet_name_list)

                ;; worksheets
                (let ([worksheets_dir (build-path xl_dir "worksheets")])
                  ;; _rels
                  (let ([worksheets_rels_dir (build-path worksheets_dir "_rels")])
                    (make-directory* worksheets_rels_dir)
                    (write-worksheets-rels-file worksheets_rels_dir sheet_count)))
                )
          )))
        (lambda () (void)))))
