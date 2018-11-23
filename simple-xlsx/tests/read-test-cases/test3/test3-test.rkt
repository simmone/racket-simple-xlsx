#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test3.xlsx")

(define test-test3
  (test-suite
   "test-test3"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"
       
       (load-sheet "第二页" xlsx)
       (check-equal? (get-cell-value "A1" xlsx) "chenxiao")
       (check-equal? (get-cell-value "B1" xlsx) 121313.23)
       (check-equal? (get-cell-value "C1" xlsx) "123456")

       ;; date type can't deal correctly, ignore
       (check-equal? (get-cell-value "D1" xlsx) 41640)
       (check-equal? (get-cell-value "D2" xlsx) 36525)
       (check-equal? (get-cell-value "D3" xlsx) 36526)
       (check-equal? (get-cell-value "D4" xlsx) 36585)
       (check-equal? (get-cell-value "D5" xlsx) 36586)

       (check-equal? (get-cell-value "E1" xlsx) 100)
       (check-equal? (get-cell-value "F1" xlsx) 200)
       (check-equal? (get-cell-value "G1" xlsx) 300)

       (check-equal? (get-cell-value "H1" xlsx) -23423423.34)

       (check-equal? (get-cell-value "I1" xlsx) 2343.3)

       )

      ))))

(run-tests test-test3)
