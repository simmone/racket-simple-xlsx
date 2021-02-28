#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader/load-workbook.rkt")

(require "../../../xlsx/xlsx.rkt")

(require racket/runtime-path)
(define-runtime-path workbook_file "workbook.xml")

(define test-load-workbook
  (test-suite
   "test-load-workbook"
   
   (test-case
    "test-load-workbook"

    (let ([_xlsx (new-xlsx)])
      (load-workbook workbook_file _xlsx)

      (check-equal? (xlsx-sheet_count _xlsx) 10)

      (check-equal? (hash-ref (xlsx-sheet_index_id_map _xlsx) 0) "1")
      (check-equal? (hash-ref (xlsx-sheet_index_name_map _xlsx) 0) "DataSheet")
      (check-equal? (hash-ref (xlsx-sheet_name_index_map _xlsx) "DataSheet") 0)
      (check-equal? (hash-ref (xlsx-sheet_index_rid_map _xlsx) 0) "rId1")

      (check-equal? (hash-ref (xlsx-sheet_index_id_map _xlsx) 4) "5")
      (check-equal? (hash-ref (xlsx-sheet_index_name_map _xlsx) 4) "LineChart2")
      (check-equal? (hash-ref (xlsx-sheet_name_index_map _xlsx) "LineChart2") 4)
      (check-equal? (hash-ref (xlsx-sheet_index_rid_map _xlsx) 4) "rId5")

      (check-equal? (hash-ref (xlsx-sheet_index_id_map _xlsx) 9) "10")
      (check-equal? (hash-ref (xlsx-sheet_index_name_map _xlsx) 9) "PieChart3D")
      (check-equal? (hash-ref (xlsx-sheet_name_index_map _xlsx) "PieChart3D") 9)
      (check-equal? (hash-ref (xlsx-sheet_index_rid_map _xlsx) 9) "rId10")

      ))))

(run-tests test-load-workbook)
