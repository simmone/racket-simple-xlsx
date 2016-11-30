#lang racket

(provide (contract-out
          [write-xlsx-file (-> (is-a?/c xlsx%) path-string? void?)]
          ))

(require racket/date)

(require "xlsx.rkt")
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

(define (write-xlsx-file xlsx xlsx_file_name)
  (when (file-exists? xlsx_file_name)
        (delete-file xlsx_file_name))
  
  (let ([tmp_dir #f])
    (dynamic-wind
        (lambda () (set! tmp_dir (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
        (lambda ()
          ;; [Content_Types].xml
          (write-content-type-file tmp_dir (get-field sheets xlsx))

          ;; _rels
          (let ([rels_dir (build-path tmp_dir "_rels")])
            (make-directory* rels_dir)
            (write-rels-file rels_dir))

          ;; docProps
          (let ([doc_props_dir (build-path tmp_dir "docProps")])
            (make-directory* doc_props_dir)
            (write-docprops-app-file doc_props_dir (get-field sheets xlsx))
            (write-docprops-core-file doc_props_dir (current-date)))
                
          ;; xl
          (let ([xl_dir (build-path tmp_dir "xl")])
            ;; _rels
            (let ([rels_dir (build-path xl_dir "_rels")])
              (make-directory* rels_dir)
              (write-workbook-xml-rels-file rels_dir (get-field sheets xlsx)))
                  
            ;; printerSettings
            (let ([printer_settings_dir (build-path xl_dir "printerSettings")])
              (make-directory* printer_settings_dir)
              (create-printer-settings printer_settings_dir (get-field sheets xlsx)))

            ;; theme
            (let ([theme_dir (build-path xl_dir "theme")])
              (make-directory* theme_dir)
              (write-theme-file theme_dir))

            ;; sharedStrings
            (write-shared-strings-file xl_dir (get-field string_item_map xlsx))

            ;; styles
            (write-styles-file xl_dir (send xlsx get-styles-list))

            ;; workbook
            (write-workbook-file xl_dir (get-field sheets xlsx))

            ;; data-sheets
            (let ([worksheets_dir (build-path xl_dir "worksheets")])
              ;; _rels
              (let ([worksheets_rels_dir (build-path worksheets_dir "_rels")])
                (make-directory* worksheets_rels_dir)
                (write-worksheets-rels-file worksheets_rels_dir (get-field sheets xlsx)))

              ;; worksheet
              (write-data-sheet-file worksheets_dir (get-field sheets xlsx) (send xlsx get-styles-list) (get-field string_item_map xlsx))
              )


            )
          ))
    (zip-xlsx xlsx_file_name tmp_dir)))
          (lambda ()
            (delete-directory/files tmp_dir)))))
