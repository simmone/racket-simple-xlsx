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
    
   (test-case
    "test-get-string-index-map"
    (let ([index_map (get-string-index-map '(("1" "2") (1) () ("1" "2" "3" 4)))])
      (check-equal? (hash-ref index_map "1") 0)
      (check-equal? (hash-ref index_map "2") 1)
      (check-equal? (hash-ref index_map "3") 2)))
    ))

(run-tests test-worksheet)
