#lang racket

(require racket/date)

(provide (contract-out
          [read-xlsx (-> path-string? procedure? any)]
          [read-batch (-> void?)]
          ))

(require "xlsx/xlsx.rkt")
(require "style/style.rkt")
(require "style/border-style.rkt")
(require "style/font-style.rkt")
(require "style/alignment-style.rkt")
(require "style/number-style.rkt")
(require "style/fill-style.rkt")
(require "style/lib.rkt")
(require "style/sort-styles.rkt")
(require "style/set-styles.rkt")
(require "sheet/sheet.rkt")
(require "lib/lib.rkt")
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

(define (read-batch)
  ;; [Content_Types].xml
  (read-content-type)

  ;; docProps/app.xml
  (read-docprops-app)

  ;; xl/sharedStrings.xml
  (read-shared-strings)

  ;; xl/styles.xml
  (read-styles)

  ;; xl/worksheets
  (read-worksheets)

  ;; xl/charts
  (read-charts)
  )

(define (read-xlsx xlsx_file_name user_proc)
  (when (not (file-exists? xlsx_file_name))
    (error (format "xlsx file [~a] not exists!" xlsx_file_name)))

  (with-xlsx
   (lambda ()
     (dynamic-wind
         (lambda () (set-XLSX-xlsx_dir! (*XLSX*) (make-temporary-file "xlsx_read_tmp_~a" 'directory ".")))
         (lambda ()
           (unzip-xlsx xlsx_file_name (XLSX-xlsx_dir (*XLSX*)))

           (read-batch)

           (user_proc)
           )
         (lambda ()
           (delete-directory/files (XLSX-xlsx_dir (*XLSX*)))
           )))))
