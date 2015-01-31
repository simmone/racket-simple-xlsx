#lang racket

(require rackunit/text-ui)

(require rackunit "worksheet.rkt")

(define test-worksheet
  (test-suite
   "test-worksheet"
   
   (test-case
    "test-get-dimension"
    (check-equal? (get-dimension '(("1" "2") ("1") () ("1" "2" "3" "4")))
                  "D4")
                  )
    
    )

   )

(run-tests test-worksheet)
