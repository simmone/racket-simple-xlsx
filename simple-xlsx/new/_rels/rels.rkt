#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/list)
(require racket/contract)

(require "../../xlsx/xlsx.rkt")

(provide (contract-out
          [write-rels (-> string?)]
          [write-rels-file (-> void?)]
          [read-rels (-> void?)]
          ))

(define S string-append)
 
(define (write-rels) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
})

(define (write-rels-file)
  (let ([rels_dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)))])
    (make-directory* rels_dir)

    (with-output-to-file (build-path rels_dir ".rels")
      #:exists 'replace
      (lambda ()
        (printf "~a" (write-rels))))))

(define (read-rels)
  (void))
