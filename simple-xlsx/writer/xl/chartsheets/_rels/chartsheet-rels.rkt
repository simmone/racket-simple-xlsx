#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-chart-sheet-rels (-> exact-nonnegative-integer? string?)]
          [write-chart-sheet-rels-file (-> path-string? exact-nonnegative-integer? void?)]
          ))

(define S string-append)

(define (write-chart-sheet-rels typeSeq) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId@|(number->string typeSeq)|" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing@|(number->string typeSeq)|.xml"/></Relationships>
})

(define (write-chart-sheet-rels-file dir typeSeq)
  (with-output-to-file (build-path dir (format "sheet~a.xml.rels" typeSeq))
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-chart-sheet-rels typeSeq)))))

