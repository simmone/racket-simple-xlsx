#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../../xlsx/xlsx.rkt")
(require "../../../xlsx/sheet.rkt")

(provide (contract-out
          [write-workbook-xml-rels (-> (is-a?/c xlsx%) string?)]
          [write-workbook-xml-rels-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define S string-append)
 
(define (write-workbook-xml-rels xlsx) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">@|(with-output-to-string
    (lambda ()
      (let loop ([loop_list (get-field sheets xlsx)])
        (when (not (null? loop_list))
              (let ([sheet (car loop_list)])
                (if (eq? (sheet-type sheet) 'data)
                    (printf 
                     "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet~a.xml\"/>"
                     (sheet-seq sheet) (sheet-typeSeq sheet))
                    (printf 
                     "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet~a.xml\"/>"
                     (sheet-seq sheet) (sheet-typeSeq sheet))))
              (loop (cdr loop_list))))

          (let ([seq (add1 (length (get-field sheets xlsx)))])
            (printf 
              "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
              seq)
            (set! seq (add1 seq))
            (printf
              "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
              seq)
            (set! seq (add1 seq))
            (when (> (hash-count (get-field string_item_map xlsx)) 0)
              (printf
                "<Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
                seq)))))|</Relationships>
})

(define (write-workbook-xml-rels-file dir xlsx)
  (make-directory* dir)

  (with-output-to-file (build-path dir "workbook.xml.rels")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-workbook-xml-rels xlsx)))))

