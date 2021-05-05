#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../writer.rkt")
(require "../../../../lib/lib.rkt")

(require rackunit "../../../../new/xl/_rels/workbook-xml-rels.rkt")

(require racket/runtime-path)
(define-runtime-path rels1_file "workbook.xml.rels1")

(define test-workbook-xml-rels
  (test-suite
   "test-workbook-xml-rels"

   (test-case
    "test-workbook-xml-rels1"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (add-data-sheet "Sheet1" '((1)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))

      (call-with-input-file rels1_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (xl-rels))
           (lambda (actual)
             (check-lines? expected actual)))))))
   ))

(run-tests test-workbook-xml-rels)
