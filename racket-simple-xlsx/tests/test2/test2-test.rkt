#lang racket

(provide test-test2)

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(require "../../lib/lib.rkt")

(define test-test2
  (test-suite
   "test-test2"

   (with-input-from-xlsx-file
    (build-path "test2" "test2.xlsx")
    (lambda ()
      (test-case
       "test-get-sheet-data"
       
       (load-sheet "Sheet1")
       (check-equal? (get-cell-value "A1") "chenxiao ")
       (check-equal? (get-cell-value "B1") "")
       (check-equal? (get-cell-value "C1") "xiaomin")
       (check-equal? (get-cell-value "D1") "")
       (check-equal? (get-cell-value "A2") "")
       (check-equal? (get-cell-value "B2") "")
       (check-equal? (get-cell-value "C2") "")
       (check-equal? (get-cell-value "D2") "haha")
       (check-equal? (get-cell-value "A3") "")
       (check-equal? (get-cell-value "B3") "")
       (check-equal? (get-cell-value "C3") "")
       (check-equal? (get-cell-value "D3") "")
       (check-equal? (get-cell-value "A4") "陈晓")
       (check-equal? (get-cell-value "B4") "爱")
       (check-equal? (get-cell-value "C4") "陈思衡")
       (check-equal? (get-cell-value "D4") "")
       )

      (test-case
       "test-get-sheet-dimension"

       (let ([dimension (get-sheet-dimension)])
         (check-equal? (car dimension) 4)
         (check-equal? (cdr dimension) 4))
       )

      (test-case
       "test-with-row"
       (let ([row_index 1])
         (with-row
          (lambda (row)
            (when (= row_index 1)
                (check-equal? (list-ref row 0) "chenxiao ")
                (check-equal? (list-ref row 1) "")
                (check-equal? (list-ref row 2) "xiaomin")
                (check-equal? (list-ref row 3) ""))

            (when (= row_index 2)
                (check-equal? (list-ref row 0) "")
                (check-equal? (list-ref row 1) "")
                (check-equal? (list-ref row 2) "")
                (check-equal? (list-ref row 3) "haha"))

            (when (= row_index 3)
                (check-equal? (list-ref row 0) "")
                (check-equal? (list-ref row 1) "")
                (check-equal? (list-ref row 2) "")
                (check-equal? (list-ref row 3) ""))

            (when (= row_index 4)
                (check-equal? (list-ref row 0) "陈晓")
                (check-equal? (list-ref row 1) "爱")
                (check-equal? (list-ref row 2) "陈思衡")
                (check-equal? (list-ref row 3) ""))
         (set! row_index (add1 row_index))))))
      ))))

