#lang racket

(require rackunit/text-ui)

(require rackunit "../../xlsx/xlsx-lib.rkt")

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

   ))

(run-tests test-xlsx)
