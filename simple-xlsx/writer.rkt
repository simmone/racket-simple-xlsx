#lang racket

(require racket/date)

(provide (contract-out
          [write-xlsx (-> path-string? procedure? any)]
          [write-batch (-> void?)]
          ))

(require "xlsx/xlsx.rkt")
(require "style/style.rkt")
(require "style/styles.rkt")
(require "style/border-style.rkt")
(require "style/font-style.rkt")
(require "style/alignment-style.rkt")
(require "style/number-style.rkt")
(require "style/fill-style.rkt")
(require "style/assemble-styles.rkt")
(require "style/set-styles.rkt")
(require "sheet/sheet.rkt")
(require "lib/lib.rkt")
(require "lib/sheet-lib.rkt")
(require "lib/dimension.rkt")
(require "content-type.rkt")
(require "_rels/rels.rkt")
(require "docProps/docprops-app.rkt")
(require "docProps/docprops-core.rkt")
(require "xl/_rels/workbook-xml-rels.rkt")
(require "xl/printerSettings/printerSettings.rkt")
(require "xl/theme/theme.rkt")
(require "xl/sharedStrings.rkt")
(require "xl/styles/styles.rkt")
(require "xl/workbook.rkt")
(require "xl/worksheets/_rels/worksheets-rels.rkt")
(require "xl/worksheets/worksheet.rkt")
(require "xl/chartsheets/_rels/chartsheets-rels.rkt")
(require "xl/chartsheets/chartsheet.rkt")
(require "xl/charts/charts.rkt")
(require "xl/drawings/_rels/drawing-rels.rkt")
(require "xl/drawings/drawing.rkt")

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
