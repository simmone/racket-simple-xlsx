#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../writer.rkt")
(require "../../../../lib/lib.rkt")

(require rackunit "../../../../new/xl/_rels/workbook-xml-rels.rkt")

(require racket/runtime-path)
(define-runtime-path rels1_file "workbook.xml.rels1")
(define-runtime-path rels2_file "workbook.xml.rels2")

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

   (test-case
    "test-workbook-xml-rels2"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (add-data-sheet "Sheet1" '((1)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      (add-chart-sheet "Chart1" 'LINE "Chart1")
      (add-chart-sheet "Chart2" 'LINE "Chart2")
      (add-chart-sheet "Chart3" 'LINE "Chart3")

      (call-with-input-file rels2_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (xl-rels))
           (lambda (actual)
             (check-lines? expected actual)))))))

   (test-case
    "test-read-workbook-rels"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (read-workbook-rels-file rels2_file)

      (check-equal? (hash-count (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*))) 8)
      (check-equal? (hash-count (XLSX-sheet_index_rel_map (*CURRENT_XLSX*))) 8)

      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId1") "worksheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 0) "worksheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId4") "chartsheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 3) "chartsheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId5") "chartsheets/sheet2.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 4) "chartsheets/sheet2.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) "rId6") "chartsheets/sheet3.xml")
      (check-equal? (hash-ref (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) 5) "chartsheets/sheet3.xml")
      ))
   ))

(run-tests test-workbook-xml-rels)
