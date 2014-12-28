#lang racket

(require rackunit/text-ui)

(require rackunit "../../../main.rkt")

(require "../../../lib/lib.rkt")

(define test-test1
  (test-suite
   "test-test1"

   (with-input-from-xlsx-file
    "test1.xlsx"
    (lambda ()
      (test-case 
       "test-get-sheets"
       
       (check-equal? (get-sheet-names) '("Sheet1" "Sheet2" "Sheet3")))
   
      (test-case
       "test-get-sheet-data"
       
       (load-sheet "Sheet1")
       (check-equal? (get-cell-value "A1") "chenxiao")
       (check-equal? (get-cell-value "B1") "love")
       (check-equal? (get-cell-value "C1") "chensiheng")
       (check-equal? (get-cell-value "C2") 7)
       )

      (test-case
       "test-get-sheet-dimension"

       (let ([dimension (get-sheet-dimension)])
         (check-equal? (car dimension) 2)
         (check-equal? (cdr dimension) 3))
       )

      (test-case
       "test-with-row"
       (let ([row_index 1])
         (with-row
          (lambda (row)
            (when (= row_index 1)
                (check-equal? (first row) "chenxiao")
                (check-equal? (second row) "love")
                (check-equal? (third row) "chensiheng"))
            (when (= row_index 2)
                (check-equal? (first row) 2)
                (check-equal? (second row) 5)
                (check-equal? (third row) 7))
            
            (set! row_index (add1 row_index))
            ))))

      ))))

(run-tests test-test1)
