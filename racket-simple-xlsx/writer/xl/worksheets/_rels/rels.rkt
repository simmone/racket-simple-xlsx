#lang racket

(provide (contract-out
          [create-worksheets-rels (-> exact-nonnegative-integer? void?)]
          ))

(define (create-worksheets-rels sheet_nums)
  (make-directory* (build-path "xl" "worksheets" "_rels"))

  (let loop ([num 1])
    (when (<= num sheet_nums)
          (with-output-to-file (build-path "xl" "worksheets" "_rels" (string-append "sheet" (number->string num) ".xml.rels"))
            (lambda ()
              (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
              (printf "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" Target=\"../printerSettings/printerSettings~a.bin\"/></Relationships>"
                      num num)
              (loop (add1 num)))))))
