#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader/load-sheet.rkt")

(require racket/runtime-path)
(define-runtime-path sheet_file "sheet.xml")

(define test-load-sheet
  (test-suite
   "test-load-sheet"
   
   (test-case
    "test-load-sheet"

    (let-values ([(sheet_id_list sheet_id_name_map sheet_name_id_map sheet_id_rid_map)
                  (load-sheet sheet_file)])
      (check-equal? (length sheet_id_list) 10)

      (check-equal? (first sheet_id_list) "1")
      (check-equal? (hash-ref sheet_id_name_map "1") "DataSheet")
      (check-equal? (hash-ref sheet_name_id_map "DataSheet") "1")
      (check-equal? (hash-ref sheet_id_rid_map "1") "rId1")

      (check-equal? (fifth sheet_id_list) "5")
      (check-equal? (hash-ref sheet_id_name_map "5") "LineChart2")
      (check-equal? (hash-ref sheet_name_id_map "LineChart2") "5")
      (check-equal? (hash-ref sheet_id_rid_map "5") "rId5")

      (check-equal? (tenth sheet_id_list) "10")
      (check-equal? (hash-ref sheet_id_name_map "10") "PieChart3D")
      (check-equal? (hash-ref sheet_name_id_map "PieChart3D") "10")
      (check-equal? (hash-ref sheet_id_rid_map "10") "rId10")

      ))))

(run-tests test-load-sheet)
