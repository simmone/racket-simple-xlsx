#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-rels (-> string?)]
          [write-rels-file (-> path-string? void?)]
          ))

(define S string-append)
 
(define (write-rels) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
})

(define (write-rels-file dir)
  (make-directory* dir)

  (with-output-to-file (build-path dir ".rels")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-rels)))))
