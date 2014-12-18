#lang racket

(provide test-lib)

(require rackunit/text-ui)

(require rackunit "../../lib/lib.rkt")

(define test-lib
  (test-suite
   "test-lib"

   (test-case 
    "test-AZ-NUMBER"
    (check-equal? (abc->number "A") 1)
    (check-equal? (abc->number "Z") 26)
    (check-equal? (abc->number "AA") 27)
    (check-equal? (abc->number "AB") 28)
    (check-equal? (abc->number "AZ") 52)
    (check-equal? (abc->number "BA") 53)
    (check-equal? (abc->number "YZ") 676)
    (check-equal? (abc->number "ZA") 677)
    (check-equal? (abc->number "ZZ") 702)
    (check-equal? (abc->number "AAA") 703)
   )

   (test-case 
    "test-NUMBER-AZ"
    (check-equal? (number->abc 1) "A")
    (check-equal? (number->abc 26) "Z")
    (check-equal? (number->abc 27) "AA")
    (check-equal? (number->abc 28) "AB")
    (check-equal? (number->abc 29) "AC")
    (check-equal? (number->abc 51) "AY")
    (check-equal? (number->abc 52) "AZ")
    (check-equal? (number->abc 53) "BA")
    (check-equal? (number->abc 676) "YZ")
    (check-equal? (number->abc 677) "ZA")
    (check-equal? (number->abc 702) "ZZ")
    (check-equal? (number->abc 703) "AAA")
   )

   (test-case 
    "test-number->list"
    (check-equal? (number->list 1) '(1))
    (check-equal? (number->list 2) '(1 2))
    (check-equal? (number->list 5) '(1 2 3 4 5)))
   
   (test-case
    "test-find-date"
    (when (not (equal? (system-type 'os) 'windows))
          (let ([the_date (find-date 1)])
            (check-equal? (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)) "1900-1-1"))

          (let ([the_date (find-date 2)])
            (check-equal? (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)) "1900-1-2"))

          (let ([the_date (find-date 59)])
            (check-equal? (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)) "1900-2-28"))

          (let ([the_date (find-date 60)])
            (check-equal? (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)) "1900-3-1"))

          (let ([the_date (find-date 61)])
            (check-equal? (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)) "1900-3-1"))

          (let ([the_date (find-date 62)])
            (check-equal? (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)) "1900-3-2")))

    (let ([the_date (find-date 41640)])
      (check-equal? (format "~a-~a-~a" (date-year the_date) (date-month the_date) (date-day the_date)) "2014-1-1"))
    )
   
   (test-case
    "test-format-hour"
    (check-equal? (format-time 0.5560763888888889) "13:20:45")
    (check-equal? (format-time 0.2560763888888889) "06:08:45")
    (check-equal? (format-time 0.5) "12:00:00"))

   ))
