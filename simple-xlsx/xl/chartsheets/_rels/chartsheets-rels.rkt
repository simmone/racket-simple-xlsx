#lang racket

(require simple-xml)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")

(provide (contract-out
          [chartsheets-rels (-> natural? list?)]
          [write-chartsheets-rels (->* () (path-string?) void?)]
          ))

(define (chartsheets-rels typeSeq)
  (list
   "Relationships"
   (cons "xmlns" "http://schemas.openxmlformats.org/package/2006/relationships")
   (list
    "Relationship"
    (cons "Id" "rId1")
    (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
    (cons "Target" (format "../drawings/drawing~a.xml" typeSeq)))))

(define (write-chartsheets-rels [output_dir #f])
  (let ([charts (filter CHART-SHEET? (XLSX-sheet_list (*XLSX*)))])
    (when (not (null? charts))
          (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "chartsheets" "_rels"))])
            (make-directory* dir)

            (let loop ([chart_sheets charts]
                       [chart_sheet_index 1])
              (when (not (null? chart_sheets))
                    (with-output-to-file (build-path dir (format "sheet~a.xml.rels" chart_sheet_index))
                      #:exists 'replace
                      (lambda ()
                        (printf (lists->xml (chartsheets-rels chart_sheet_index)))))
                    (loop (cdr chart_sheets) (add1 chart_sheet_index))))))))
