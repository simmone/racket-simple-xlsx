#lang racket

(require rackunit/text-ui)

(require rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test6.xlsx")

(define test-test6
  (test-suite
   "test-test6"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"

       (load-sheet "Sheet1" xlsx)
;       (check-equal? (get-cell-value "A1" xlsx) "1801450575700400010001")
       (check-equal? (get-cell-value "A2" xlsx) "1801450575700400020001")
;       (check-equal? (get-cell-value "A5" xlsx) "11014564926001")
;       (check-equal? (get-cell-value "A10" xlsx) "13009922869601D0046")
;       (check-equal? (get-cell-value "A45" xlsx) "1801342996690100380001")
;       (check-equal? (get-cell-value "D45" xlsx) 100)
;       (check-equal? (get-cell-value "D14" xlsx) 81.045)
;       (check-equal? (get-cell-value "F14" xlsx) 19.001)
;       (check-equal? (get-cell-value "A25" xlsx) "11014494971201")
;       (check-equal? (get-cell-value "A44" xlsx) "1801342996690100370001        ")
      )))))

(run-tests test-test6)
