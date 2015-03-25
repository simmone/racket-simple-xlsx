#lang racket

(provide (contract-out
          [write-worksheets-rels-file (-> path-string? exact-nonnegative-integer? void?)]
          ))

(define (write-worksheets-rels-file dir sheet_nums)
  (let loop ([num 1])
    (when (<= num sheet_nums)
          (with-output-to-file (build-path dir (string-append "sheet" (number->string num) ".xml.rels"))
            (lambda ()
              (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
              (printf "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" Target=\"../printerSettings/printerSettings~a.bin\"/></Relationships>"
                      num num)
              (loop (add1 num)))))))
