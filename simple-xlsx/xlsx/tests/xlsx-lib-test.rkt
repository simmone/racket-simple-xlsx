#lang racket

(require rackunit/text-ui)

(require rackunit "../xlsx-lib.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"
   
   (test-case
    "test-check-data-list"
    
    (check-exn exn:fail? (lambda () (check-equal? (check-data-list '()) #f)))
    (check-exn exn:fail? (lambda () ((check-data-list '((1) 4)))))
    (check-exn exn:fail? (lambda () (check-data-list '((1) (1 2)))))
    
    (check-true (check-data-list '((1 2) (3 4))))
    )

   (test-case
    "test-convert-range"
    
    (check-equal? (convert-range "C2-C10") "$C$2:$C$10")

    (check-equal? (convert-range "C2-Z2") "$C$2:$Z$2")

    (check-equal? (convert-range "AB20-AB100") "$AB$20:$AB$100")
    )
   
   (test-case
    "test-check-range"
    
    (check-true (check-range "A2-Z2"))

    (check-exn exn:fail? (lambda () (check-true (check-range "A2-C3"))))
    
    (check-exn exn:fail? (lambda () (check-range "c2")))
    (check-exn exn:fail? (lambda () (check-range "c2-c2")))

    (check-exn exn:fail? (lambda () (check-range "A2-A1")))
    (check-exn exn:fail? (lambda () (check-range "A2-B3")))
    )
   
   (test-case
    "test-check-col-range"
    
    (check-equal? (check-col-range "A-Z") "A-Z")

    (check-equal? (check-col-range "A") "A-A")

    (check-equal? (check-col-range "10") "10-10")
    
    (check-exn exn:fail? (lambda () (check-col-range "B-A")))

    (check-exn exn:fail? (lambda () (check-col-range "A1-A")))
    )

   (test-case
    "test-check-cell-range"
    
    (check-cell-range "A1-B2")

    (check-exn exn:fail? (lambda () (check-cell-range "A10-B9")))

    (check-exn exn:fail? (lambda () (check-cell-range "B1-A1")))
    )

   (test-case
    "test-range-length"
    
    (check-equal? (range-length "A2-A20") 19)
    (check-equal? (range-length "AB21-AB21") 1)
    (check-equal? (range-length "A2-D2") 4)
    )

   ))

(run-tests test-xlsx)
