#lang racket

(require simple-xml)

(require racket/file)
(require racket/port)
(require racket/list)
(require racket/contract)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")
(require "../../lib/lib.rkt")

(provide (contract-out
          [write-docprops-core (-> date? string?)]
          [write-docprops-core-file (-> date? void?)]
          ))

(define S string-append)
 
(define (write-docprops-core the_date) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2006-09-13T11:21:51Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">@|(format-w3cdtf the_date)|</dcterms:modified></cp:coreProperties>
})

(define (write-docprops-core-file the_date)
  (let ([docprops_dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "docProps")])
    (make-directory* docprops_dir)

    (with-output-to-file (build-path docprops_dir "core.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (write-docprops-core the_date))))))
