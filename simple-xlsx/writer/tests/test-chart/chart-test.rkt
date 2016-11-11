#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "chart.xlsx")

(define test-chart
  (test-suite
   "test-chart"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"
       
       (load-sheet "Sheet1" xlsx)
       (check-equal? (get-cell-value "A2" xlsx) 201601)
       )

      ))))

(run-tests test-chart)
