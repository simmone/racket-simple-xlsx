#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(require "../../../define.rkt")

(provide (contract-out
          [write-workbook-xml-rels (-> list? string?)]
          [write-workbook-xml-rels-file (-> path-string? list? void?)]
          ))

(define S string-append)
 
(define (write-workbook-xml-rels sheet_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">@|(let ([seq 1])
(with-output-to-string
    (lambda ()
      (let loop ([loop_list sheet_list])
        (when (not (null? loop_list))
              (let ([sheet (car loop_list)])
                (if (eq? (sheetData-type sheet) 'data)
                    (printf 
                     "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet~a.xml\"/>"
                     (sheetData-seq sheet) (sheetData-typeSeq sheet))
                    (printf 
                     "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet~a.xml\"/>"
                     (sheetData-seq sheet) (sheetData-typeSeq sheet))))
              (loop (cdr loop_list))))

          (set! seq (add1 (length sheet_list)))
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
            seq))))|</Relationships>
})

(define (write-workbook-xml-rels-file dir sheet_list)
  (with-output-to-file (build-path dir "workbook.xml.rels")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-workbook-xml-rels sheet_list)))))

