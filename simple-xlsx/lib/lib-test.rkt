#lang racket

(require rackunit/text-ui)
(require racket/date)

(require rackunit "lib.rkt")

(define test-lib
  (test-suite
   "test-lib"

   (test-case 
    "test-AZ-NUMBER"
    (check-equal? (abc->number "A") 1)
    (check-equal? (abc->number "B") 2)
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
    "test-AZ-RANGE"
    (check-equal? (abc->range "A") '(1 . 1))
    (check-equal? (abc->range "B") '(2 . 2))
    (check-equal? (abc->range "2") '(2 . 2))
    (check-equal? (abc->range "C-D") '(3 . 4))
    (check-equal? (abc->range "3-4") '(3 . 4))
    (check-equal? (abc->range "1-26") '(1 . 26))
    (check-equal? (abc->range "26-1") '(1 . 1))
    (check-equal? (abc->range "A-Z") '(1 . 26))
    (check-equal? (abc->range "Z-A") '(1 . 1))
    (check-equal? (abc->range "A-B-C") '(1 . 1))
    (check-equal? (abc->range "A-ksdk344") '(1 . 1))
    (check-equal? (abc->range "sdksjdkf-%^%$#") '(1 . 1))
    )

   (test-case
    "test-get-range-ref"
    (let ([range_hash (make-hash)])
      (hash-set! range_hash '(1 . 2) 1)
      (hash-set! range_hash '(3 . 4) 2)
      
      (check-equal? (get-range-ref range_hash 1) 1) 
      (check-equal? (get-range-ref range_hash 2) 1)
      (check-equal? (get-range-ref range_hash 3) 2)
      (check-equal? (get-range-ref range_hash 4) 2)
      (check-equal? (get-range-ref range_hash 5) #f)
      (check-equal? (get-range-ref range_hash 0) #f)
      )
    )

   (test-case 
    "test-number->list"
    (check-equal? (number->list 1) '(1))
    (check-equal? (number->list 2) '(1 2))
    (check-equal? (number->list 5) '(1 2 3 4 5)))
   
   (test-case
    "test-format-hour"
    (check-equal? (format-time 0.5560763888888889) "13:20:45")
    (check-equal? (format-time 0.2560763888888889) "06:08:45")
    (check-equal? (format-time 0.5) "12:00:00"))
   
   (test-case
    "test-value-of-time"
    (check-equal? (format-complete-time (value-of-time "20141231 12:31:23")) "20141231 12:31:23")
    (check-equal? (format-complete-time (value-of-time "20140105 12:31:23")) "20140105 12:31:23"))

   (test-case
    "test-format-w3cdtf"
    (check-equal? (format-w3cdtf (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")) "2015-01-02T13:17:44+08:00"))
   
   (test-case
    "test-create-sheet-name-list"
    (check-equal? (create-sheet-name-list 5) '("Sheet1" "Sheet2" "Sheet3" "Sheet4" "Sheet5"))
    (check-equal? (create-sheet-name-list 1) '("Sheet1")))

   (test-case
    "test-get-dimension"
    (check-equal? (get-dimension '(("1" "2") ("1") () ("1" "2" "3" "4")))
                  "D4")
                  )
    
   (test-case
    "test-get-string-index-map"
    (let-values ([(index_list index_map) (get-string-index '((("1" "2") (1) () ("1" "2" "3" 4)) (("7" "8"))))])
      (check-equal? (hash-ref index_map "1") 0)
      (check-equal? (hash-ref index_map "2") 1)
      (check-equal? (hash-ref index_map "3") 2)
      (check-equal? (hash-ref index_map "7") 3)
      (check-equal? (hash-ref index_map "8") 4)

      (check-equal? (list-ref index_list 0) "1")
      (check-equal? (list-ref index_list 1) "2")
      (check-equal? (list-ref index_list 2) "3")
      (check-equal? (list-ref index_list 3) "7")
      (check-equal? (list-ref index_list 4) "8")
      ))
   
    ))

(run-tests test-lib)
