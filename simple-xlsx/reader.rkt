#lang racket

(provide (contract-out
          [read-xlsx (-> path-string? procedure? any)]
          [read-batch (-> void?)]
          ))

(require racket/date
         "xlsx/xlsx.rkt"
         "style/style.rkt"
         "style/border-style.rkt"
         "style/font-style.rkt"
         "style/alignment-style.rkt"
         "style/number-style.rkt"
         "style/fill-style.rkt"
         "style/set-styles.rkt"
         "sheet/sheet.rkt"
         "lib/lib.rkt"
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

(define (read-batch)
  ;; [Content_Types].xml
  (read-content-type)

  ;; docProps/app.xml
  (read-docprops-app)

  ;; xl/workbook.xml
  (read-workbook)

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
