#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader/load-workbook-rels.rkt")

(require "../../../xlsx/xlsx.rkt")

(require racket/runtime-path)
(define-runtime-path workbook_rels_file "workbook.xml.rels")

(define test-load-workbook-rels
  (test-suite
   "test-load-workbook-rels"
   
   (test-case
    "test-load-workbook-rels"

    (let ([_xlsx (new-xlsx)])
      (load-workbook-rels workbook_rels_file _xlsx)

      (check-equal? (hash-count (XLSX-sheet_rid_rel_map _xlsx)) 13)

      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map _xlsx) "rId1") "worksheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map _xlsx) "rId4") "chartsheets/sheet1.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map _xlsx) "rId5") "chartsheets/sheet2.xml")
      (check-equal? (hash-ref (XLSX-sheet_rid_rel_map _xlsx) "rId10") "chartsheets/sheet7.xml")
      ))
    
    ))

(run-tests test-load-workbook-rels)
