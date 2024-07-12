#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt")

(provide (contract-out
          [drawing (-> natural? list?)]
          [write-drawings (->* () (path-string?) void?)]
          ))

(define (drawing type_seq)
  (list
   "xdr:wsDr"
   (cons "xmlns:xdr" "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
   (cons "xmlns:a" "http://schemas.openxmlformats.org/drawingml/2006/main")
   (list
    "xdr:absoluteAnchor"
    (list
     "xdr:pos"
     (cons "x" "0")
     (cons "y" "0"))
    (list
     "xdr:ext"
     (cons "cx" "9311355")
     (cons "cy" "6088879"))
    (list
     "xdr:graphicFrame"
     (cons "macro" "")
     (list
      "xdr:nvGraphicFramePr"
      (list "xdr:cNvPr"
            (cons "id" (format "~a" type_seq))
            (cons "name" "图表 1"))
      (list
       "xdr:cNvGraphicFramePr"
       (list
        "a:graphicFrameLocks"
        (cons "noGrp" "1"))))
     (list
      "xdr:xfrm"
      (list
       "a:off"
       (cons "x" "0")
       (cons "y" "0"))
      (list
       "a:ext"
       (cons "cx" "0")
       (cons "cy" "0")))
     (list
      "a:graphic"
      (list
       "a:graphicData"
       (cons "uri" "http://schemas.openxmlformats.org/drawingml/2006/chart")
       (list
        "c:chart"
        (cons "xmlns:c" "http://schemas.openxmlformats.org/drawingml/2006/chart")
        (cons "xmlns:r" "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        (cons "r:id" (format "rId~a" type_seq))))))
    (list
     "xdr:clientData"))))

(define (write-drawings [output_dir #f])
  (let ([charts (filter CHART-SHEET? (XLSX-sheet_list (*XLSX*)))])
    (when (not (null? charts))

          (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "drawings"))])
            (make-directory* dir)

            (let loop ([chart_sheets charts]
                       [chart_sheet_index 1])
              (when (not (null? chart_sheets))
                    (with-output-to-file (build-path dir (format "drawing~a.xml" chart_sheet_index))
                      #:exists 'replace
                      (lambda ()
                        (printf (lists-to-xml (drawing chart_sheet_index)))))
                    (loop (cdr chart_sheets) (add1 chart_sheet_index))))))))
