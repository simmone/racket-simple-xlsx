#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-drawing-rels (-> exact-nonnegative-integer? string?)]
          [write-drawing-rels-file (-> path-string? exact-nonnegative-integer? void?)]
          ))

(define S string-append)

(define (write-drawing-rels typeSeq) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId@|(number->string typeSeq)|" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart@|(number->string typeSeq)|.xml"/></Relationships>
})

(define (write-drawing-rels-file dir typeSeq)
  (with-output-to-file (build-path dir (format "drawing~a.xml.rels" typeSeq))
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-drawing-rels typeSeq)))))

