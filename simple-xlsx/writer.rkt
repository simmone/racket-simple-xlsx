#lang racket

(provide (contract-out
          [write-xlsx (-> path-string? procedure? any)]
          [write-batch (-> void?)]
          ))

(require racket/date
         "xlsx/xlsx.rkt"
         "style/style.rkt"
         "style/styles.rkt"
         "style/border-style.rkt"
         "style/font-style.rkt"
         "style/alignment-style.rkt"
         "style/number-style.rkt"
         "style/fill-style.rkt"
         "style/assemble-styles.rkt"
         "style/set-styles.rkt"
         "sheet/sheet.rkt"
         "lib/lib.rkt"
         "lib/sheet-lib.rkt"
         "lib/dimension.rkt"
         "content-type.rkt"
         "_rels/rels.rkt"
         "docProps/docprops-app.rkt"
         "docProps/docprops-core.rkt"
         "xl/_rels/workbook-xml-rels.rkt"
         "xl/printerSettings/printerSettings.rkt"
         "xl/theme/theme.rkt"
         "xl/sharedStrings.rkt"
         "xl/styles/styles.rkt"
         "xl/workbook.rkt"
         "xl/worksheets/_rels/worksheets-rels.rkt"
         "xl/worksheets/worksheet.rkt"
         "xl/chartsheets/_rels/chartsheets-rels.rkt"
         "xl/chartsheets/chartsheet.rkt"
         "xl/charts/charts.rkt"
         "xl/drawings/_rels/drawing-rels.rkt"
         "xl/drawings/drawing.rkt")

(define (write-batch)
  ;; write shared_strings_map
  (squash-shared-strings-map)

  ;; [Content_Types].xml
  (write-content-type)

  ;; _rels/.rels
  (write-rels)

  ;; docProps/app.xml
  (write-docprops-app)

  ;; docProps/core.xml
  (write-docprops-core (current-date))

  ;; xl/_rels/rels
  (write-workbook-rels)

  ;; xl/theme/theme1.xml
  (write-theme)

  ;; xl/sharedStrings.xml
  (write-shared-strings)

  ;; xl/styles.xml
  (strip-styles)
  (assemble-styles)
  (write-styles)

  ;; xl/workbook
  (write-workbook)

  ;; xl/worksheets/_rels/sheet?.rels.xml
  (write-worksheets-rels)

  ;; xl/worksheets/worksheet/sheet?.xml
  (write-worksheets)

  ;; xl/chartsheets/_rels/sheet?.rels.xml
  (write-chartsheets-rels)

  ;; xl/chartsheets/chartsheet/sheet?.xml
  (write-chartsheets)

  ;; xl/charts/chart?.xml
  (write-charts)

  ;; xl/drawings/_rels/drawing?.rels.xml
  (write-drawings-rels)

  ;; xl/drawings/drawing?.xml
  (write-drawings))

(define (write-xlsx xlsx_file_name user_proc)
  (when (file-exists? xlsx_file_name)
    (delete-file xlsx_file_name))

  (with-xlsx
   (lambda ()
     (dynamic-wind
         (lambda () (set-XLSX-xlsx_dir! (*XLSX*) (make-temporary-file "xlsx_write_tmp_~a" 'directory ".")))
         (lambda ()
           (user_proc)

           (write-batch)

           (zip-xlsx xlsx_file_name (XLSX-xlsx_dir (*XLSX*))))
         (lambda ()
;;           (void)
           (delete-directory/files (XLSX-xlsx_dir (*XLSX*)))
           )))))
