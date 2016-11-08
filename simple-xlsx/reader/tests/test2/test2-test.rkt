#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "test2.xlsx")

(define test-test2
  (test-suite
   "test-test2"

   (with-input-from-xlsx-file
    test_file
    (lambda (xlsx)
      (test-case
       "test-get-sheet-data"
       
       (load-sheet "Sheet1" xlsx)
       (check-equal? (get-cell-value "A1" xlsx) "chenxiao ")
       (check-equal? (get-cell-value "B1" xlsx) "")
       (check-equal? (get-cell-value "C1" xlsx) "xiaomin")
       (check-equal? (get-cell-value "D1" xlsx) "")
       (check-equal? (get-cell-value "A2" xlsx) "")
       (check-equal? (get-cell-value "B2" xlsx) "")
       (check-equal? (get-cell-value "C2" xlsx) "")
       (check-equal? (get-cell-value "D2" xlsx) "haha")
       (check-equal? (get-cell-value "A3" xlsx) "")
       (check-equal? (get-cell-value "B3" xlsx) "")
       (check-equal? (get-cell-value "C3" xlsx) "")
       (check-equal? (get-cell-value "D3" xlsx) "")
       (check-equal? (get-cell-value "A4" xlsx) "陈晓")
       (check-equal? (get-cell-value "B4" xlsx) "爱")
       (check-equal? (get-cell-value "C4" xlsx) "陈思衡")
       (check-equal? (get-cell-value "D4" xlsx) "")
       )

      (test-case
       "test-get-sheet-dimension"

       (let ([dimension (get-sheet-dimension xlsx)])
         (check-equal? (car dimension) 4)
         (check-equal? (cdr dimension) 4))
       )

      (test-case
       "test-with-row"
       (let ([row_index 1])
         (with-row xlsx
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

(run-tests test-test2)
