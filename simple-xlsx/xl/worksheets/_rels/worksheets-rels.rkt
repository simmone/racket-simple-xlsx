#lang racket

(require simple-xml)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")

(provide (contract-out
          [worksheets-rels (-> natural? list?)]
          [write-worksheets-rels (->* () (path-string?) void?)]
          ))

(define (worksheets-rels sheet_index)
  (append
   '("Relationships" ("xmlns" . "http://schemas.openxmlformats.org/package/2006/relationships"))
   (list
    (list "Relationship"
          (cons "Id" (format "rId~a" sheet_index))
          (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings")
          (cons "Target" (format "../printerSettings/printerSettings~a.bin" sheet_index))))))

(define (write-worksheets-rels [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "worksheets" "_rels"))])
    (make-directory* dir)

    (let loop ([data_sheets (filter DATA-SHEET? (XLSX-sheet_list (*XLSX*)))]
               [data_sheet_index 1])
      (when (not (null? data_sheets))
        (with-output-to-file (build-path dir (format "sheet~a.xml.rels" data_sheet_index))
          #:exists 'replace
          (lambda ()
            (printf (lists->xml (worksheets-rels data_sheet_index)))))
        (loop (cdr data_sheets) (add1 data_sheet_index))))))
