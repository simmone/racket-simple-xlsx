#lang racket

(require fast-xml
         "../../../xlsx/xlsx.rkt"
         "../../../sheet/sheet.rkt")

(provide (contract-out
          [drawing-rels (-> natural? list?)]
          [write-drawings-rels (->* () (path-string?) void?)]
          ))

(define (drawing-rels type_seq)
  (list
   "Relationships"
   (cons "xmlns" "http://schemas.openxmlformats.org/package/2006/relationships")
   (list
    "Relationship"
    (cons "Id" (format "rId~a" type_seq))
    (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
    (cons "Target" (format "../charts/chart~a.xml" type_seq)))))

(define (write-drawings-rels [output_dir #f])
  (let ([charts (filter CHART-SHEET? (XLSX-sheet_list (*XLSX*)))])
    (when (not (null? charts))
          (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "drawings" "_rels"))])
            (make-directory* dir)

            (let loop ([chart_sheets charts]
                       [chart_sheet_index 1])
              (when (not (null? chart_sheets))
                    (with-output-to-file (build-path dir (format "drawing~a.xml.rels" chart_sheet_index))
                      #:exists 'replace
                      (lambda ()
                        (printf (lists-to-xml (drawing-rels chart_sheet_index)))))
                    (loop (cdr chart_sheets) (add1 chart_sheet_index))))))))
