#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-workbook-xml-rels (-> exact-nonnegative-integer? string?)]
          [write-workbook-xml-rels-file (-> path-string? exact-nonnegative-integer? void?)]
          ))

(define S string-append)
 
(define (write-workbook-xml-rels sheet_count) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">@|(let ([seq 1])
(with-output-to-string
    (lambda ()
      (let loop ([count 1])
        (when (<= count sheet_count)
          (printf 
           "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet~a.xml\"/>"
           count count)
          (set! seq count)
          (loop (add1 count))))
          (set! seq (add1 seq))
          (printf 
            "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
            seq)
          (set! seq (add1 seq))
          (printf
            "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
            seq)
          (set! seq (add1 seq))
          (printf
            "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
            seq))))|
})

(define (write-workbook-xml-rels-file dir sheet_count)
  (with-output-to-file (build-path dir "workbook.xml.rels")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-workbook-xml-rels sheet_count)))))

