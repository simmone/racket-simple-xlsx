#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(require "../../../lib/lib.rkt")

(define test-test6
  (test-suite
   "test-test6"

   (with-input-from-xlsx-file
    "test6.xlsx"
    (lambda ()
      (test-case
       "test-get-sheet-data"

       (load-sheet "Sheet1")
       (check-equal? (get-cell-value "A1") "1801450575700400010001")
       (check-equal? (get-cell-value "A2") "1801450575700400020001")
       (check-equal? (get-cell-value "A5") "11014564926001")
       (check-equal? (get-cell-value "A10") "13009922869601D0046")
       (check-equal? (get-cell-value "A45") "1801342996690100380001")
       (check-equal? (get-cell-value "D45") 100)
       (check-equal? (get-cell-value "D14") 81.045)
       (check-equal? (get-cell-value "F14") 19.001)
       (check-equal? (get-cell-value "A25") "11014494971201")
       (check-equal? (get-cell-value "A44") "1801342996690100370001        ")
      )))))

(run-tests test-test6)
