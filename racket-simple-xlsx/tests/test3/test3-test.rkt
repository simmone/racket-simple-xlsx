#lang racket

(provide test-test3)

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(require "../../lib/lib.rkt")

(define test-test3
  (test-suite
   "test-test3"

   (with-input-from-excel-file
    (build-path "test3" "test3.xlsx")
    (lambda ()
      (test-case
       "test-get-sheet-data"
       
       (load-sheet "第二页")
       (check-equal? (get-cell-value "A1") "chenxiao")
       (check-equal? (get-cell-value "B1") 121313.23)
       (check-equal? (get-cell-value "C1") "123456")

       ;; date type can't deal correctly, ignore
       (check-equal? (get-cell-value "D1") 41640)
       (check-equal? (get-cell-value "D2") 36525)
       (check-equal? (get-cell-value "D3") 36526)
       (check-equal? (get-cell-value "D4") 36585)
       (check-equal? (get-cell-value "D5") 36586)

       (check-equal? (get-cell-value "E1") 100)
       (check-equal? (get-cell-value "F1") 200)
       (check-equal? (get-cell-value "G1") 300)

       (check-equal? (get-cell-value "H1") -23423423.34)

       (check-equal? (get-cell-value "I1") 2343.3)

       (check-equal? (get-cell-value "J1") "13:20:45")

       )

      ))))
