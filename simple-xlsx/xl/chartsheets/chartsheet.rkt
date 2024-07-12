#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt")

(provide (contract-out
          [chart-sheet (-> natural? list?)]
          [write-chartsheets (->* () (path-string?) void?)]
          ))

(define (chart-sheet type_seq)
  (list
   "chartsheet"
   (cons "xmlns" "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
   (cons "xmlns:r" "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
   (list "sheetPr")
   (list
    "sheetViews"
    (list "sheetView"
          (cons "zoomScale" "115")
          (cons "workbookViewId" "0")
          (cons "zoomToFit" "1")))
   (list
    "pageMargins"
    (cons "left" "0.7")
    (cons "right" "0.7")
    (cons "top" "0.75")
    (cons "bottom" "0.75")
    (cons "header" "0.3")
    (cons "footer" "0.3"))
   (list
    "drawing"
    (cons "r:id" (format "rId~a" type_seq)))))

(define (write-chartsheets [output_dir #f])
  (let ([charts (filter CHART-SHEET? (XLSX-sheet_list (*XLSX*)))])
    (when (not (null? charts))
          (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "chartsheets"))])
            (make-directory* dir)

            (let loop ([chart_sheets charts]
                       [chart_sheet_index 1])
              (when (not (null? chart_sheets))
                    (with-output-to-file (build-path dir (format "sheet~a.xml" chart_sheet_index))
                      #:exists 'replace
                      (lambda ()
                        (printf (lists-to-xml (chart-sheet chart_sheet_index)))))
                    (loop (cdr chart_sheets) (add1 chart_sheet_index))))))))
