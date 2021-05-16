#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")
(require "../../lib/lib.rkt")

(provide (contract-out
          [docprops-core (-> date? list?)]
          [write-docprops-core (-> date? void?)]
          [read-docpros-core (-> void?)]
          ))

(define (docprops-core the_date)
  (list
   "cp:coreProperties"
   (cons "xmlns:cp" "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
   (cons "xmlns:dc" "http://purl.org/dc/elements/1.1/")
   (cons "xmlns:dcterms" "http://purl.org/dc/terms/")
   (cons "xmlns:dcmitype" "http://purl.org/dc/dcmitype/")
   (cons "xmlns:xsi" "http://www.w3.org/2001/XMLSchema-instance")
   (list "dc:creator" "")
   (list "cp:lastModifiedBy" "")
   (list "dcterms:created"
         (cons "xsi:type" "dcterms:W3CDTF")
         (format-w3cdtf the_date))
   (list "dcterms:modified"
         (cons "xsi:type" "dcterms:W3CDTF")
         (format-w3cdtf the_date))))

(define (write-docprops-core the_date)
  (let ([docprops_dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "docProps")])
    (make-directory* docprops_dir)

    (with-output-to-file (build-path docprops_dir "core.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists->compact_xml (docprops-core)))))))

(define (read-docpros-core)
  (void))
