#lang racket

(require simple-xml)

(require "../xlsx/xlsx.rkt")

(provide (contract-out
          [rels (-> list?)]
          [write-rels (->* () (path-string?) void?)]
          ))

(define (rels)
  '("Relationships"
    ("xmlns" . "http://schemas.openxmlformats.org/package/2006/relationships")
    ("Relationship"
     ("Id" . "rId3")
     ("Type" . "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
     ("Target" . "docProps/app.xml"))
    ("Relationship"
     ("Id". "rId2")
     ("Type" . "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
     ("Target" . "docProps/core.xml"))
    ("Relationship"
     ("Id" . "rId1")
     ("Type" . "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
     ("Target" . "xl/workbook.xml"))))

(define (write-rels [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "_rels"))])
    (make-directory* dir)

    (with-output-to-file (build-path dir ".rels")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists->xml (rels)))))))
