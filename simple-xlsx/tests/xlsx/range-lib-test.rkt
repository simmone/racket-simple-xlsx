#lang racket

(require rackunit/text-ui)

(require rackunit "../../xlsx/range-lib.rkt")

(define test-range-lib
  (test-suite
   "test-range-lib"

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
    "test-get-dimension"
    (check-equal? (get-dimension '(("1" "2") ("1") () ("1" "2" "3" "4"))) "D4")
    (check-equal? (get-dimension '(("1" "2" "3" "4") ("3" "4" "5" "6"))) "D2")
    )
   
   (test-case
    "test-convert-range"
    
    (check-equal? (convert-range "C2-C10") "$C$2:$C$10")

    (check-equal? (convert-range "C2-Z2") "$C$2:$Z$2")

    (check-equal? (convert-range "AB20-AB100") "$AB$20:$AB$100")
    )
   
   (test-case
    "test-check-range"
    
    (check-true (only-one-row/col-data? "A2-Z2"))

    (check-exn exn:fail? (lambda () (check-true (only-one-row/col-data? "A2-C3"))))
    
    (check-exn exn:fail? (lambda () (only-one-row/col-data? "c2")))
    (check-exn exn:fail? (lambda () (only-one-row/col-data? "c2-c2")))

    (check-exn exn:fail? (lambda () (only-one-row/col-data? "A2-A1")))
    (check-exn exn:fail? (lambda () (only-one-row/col-data? "A2-B3")))
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

   (test-case
    "test-range-to-cell-hash"
    
    (let* ([range1_hash (range-to-cell-hash "A1-A2" 1)]
           [range2_hash (range-to-cell-hash "A3-B4" 2)])

      (check-equal? (hash-count range1_hash) 2)
      (check-equal? (hash-count range2_hash) 4)
      
      (check-equal? (hash-ref range1_hash "A1") 1)
      (check-equal? (hash-ref range1_hash "A2") 1)
      (check-equal? (hash-ref range2_hash "A3") 2)
      (check-equal? (hash-ref range2_hash "A4") 2)
      (check-equal? (hash-ref range2_hash "B3") 2)
      (check-equal? (hash-ref range2_hash "B4") 2))

    (let* ([range1_hash (range-to-cell-hash "A2-A1" 1)]
           [range2_hash (range-to-cell-hash "A5-E14" 2)])
      
      (check-equal? (hash-count range1_hash) 0)
      (check-equal? (hash-count range2_hash) 50)

      (check-equal? (hash-ref range2_hash "A5") 2)
      (check-equal? (hash-ref range2_hash "B10") 2)
      (check-equal? (hash-ref range2_hash "E14") 2)
      )

    )

   (test-case
    "test-range-to-row-hash"
    
    (let* ([range_hash (range-to-row-hash "1-10" 1)])

      (check-equal? (hash-count range_hash) 10)
      
      (check-equal? (hash-ref range_hash 1) 1)
      (check-equal? (hash-ref range_hash 10) 1))
    )

   (test-case
    "test-range-to-col-hash"
    
    (let* ([range_hash (range-to-col-hash "1-10" 1)])

      (check-equal? (hash-count range_hash) 10)
      
      (check-equal? (hash-ref range_hash 1) 1)
      (check-equal? (hash-ref range_hash 10) 1))

    (let* ([range_hash (range-to-col-hash "A-J" 1)])

      (check-equal? (hash-count range_hash) 10)
      
      (check-equal? (hash-ref range_hash 1) 1)
      (check-equal? (hash-ref range_hash 10) 1))

    )

   (test-case
    "test-combine-hash-in-hash"

    (let* ([range1_hash (range-to-cell-hash "A1-C3" (make-hash '((a . 1))))]
           [range2_hash (range-to-cell-hash "A3-B4" (make-hash '((a . 2) (b . 1))))]
           [range3_hash (range-to-cell-hash "B3-D5" (make-hash '((c . 3))))]
           [result_map (combine-hash-in-hash (list range1_hash range2_hash range3_hash))])
      
      (check-equal? (hash-count range1_hash) 9)
      (check-equal? (hash-count range2_hash) 4)
      (check-equal? (hash-count range3_hash) 9)
      (check-equal? (hash-count result_map) 17)
      
      (check-equal? (hash-ref result_map "A1") (make-hash '((a . 1))))
      (check-equal? (hash-ref result_map "C2") (make-hash '((a . 1))))

      (check-equal? (hash-ref result_map "A3") (make-hash '((a . 2) (b . 1))))
      (check-equal? (hash-ref result_map "A4") (make-hash '((a . 2) (b . 1))))

      (check-equal? (hash-ref result_map "B3") (make-hash '((a . 2) (b . 1) (c . 3))))
      (check-equal? (hash-ref result_map "B4") (make-hash '((a . 2) (b . 1) (c . 3))))
      (check-equal? (hash-ref result_map "C3") (make-hash '((a . 1) (c . 3))))
      (check-equal? (hash-ref result_map "D5") (make-hash '((c . 3))))

    ))

   (test-case
    "test-combine-cols-hash"
    
    (let* ([width_hash (make-hash)]
           [style_hash (make-hash)]
           [result_list #f])

      (hash-set! width_hash 1 10)
      (hash-set! width_hash 2 10)
      (hash-set! style_hash 3 3)
      (hash-set! style_hash 4 3)
      
      (set! result_list (combine-cols-hash width_hash style_hash))
      
      (check-equal? (length result_list) 2)
      (check-equal? (list-ref result_list 0) '( (1 . 2) 10 #f) )
      (check-equal? (list-ref result_list 1) '( (3 . 4) #f 3) ))

    (let* ([width_hash (make-hash)]
           [style_hash (make-hash)]
           [result_list #f])

      (hash-set! width_hash 1 10)
      (hash-set! width_hash 2 10)
      (hash-set! width_hash 3 10)
      (hash-set! style_hash 2 3)
      (hash-set! style_hash 3 3)
      (hash-set! style_hash 4 3)
      
      (set! result_list (combine-cols-hash width_hash style_hash))
      
      (check-equal? (length result_list) 3)
      (check-equal? (list-ref result_list 0) '( (1 . 1) 10 #f) )
      (check-equal? (list-ref result_list 1) '( (2 . 3) 10 3) )
      (check-equal? (list-ref result_list 2) '( (4 . 4) #f 3) ))

    (let* ([width_hash (make-hash)]
           [style_hash (make-hash)]
           [result_list #f])

      (hash-set! width_hash 1 10)
      (hash-set! width_hash 3 10)
      (hash-set! style_hash 2 3)
      (hash-set! style_hash 3 3)
      (hash-set! style_hash 4 3)
      
      (set! result_list (combine-cols-hash width_hash style_hash))
      
      (check-equal? (length result_list) 4)
      (check-equal? (list-ref result_list 0) '( (1 . 1) 10 #f) )
      (check-equal? (list-ref result_list 1) '( (2 . 2) #f 3) )
      (check-equal? (list-ref result_list 2) '( (3 . 3) 10 3) )
      (check-equal? (list-ref result_list 3) '( (4 . 4) #f 3) ))
    )

    (let* ([width_hash (make-hash)]
           [style_hash (make-hash)]
           [result_list #f])

      (hash-set! width_hash 1 10)
      (hash-set! width_hash 3 10)
      (hash-set! width_hash 5 10)
      (hash-set! style_hash 2 3)
      (hash-set! style_hash 5 3)
      
      (set! result_list (combine-cols-hash width_hash style_hash))
      
      (check-equal? (length result_list) 4)
      (check-equal? (list-ref result_list 0) '( (1 . 1) 10 #f) )
      (check-equal? (list-ref result_list 1) '( (2 . 2) #f 3) )
      (check-equal? (list-ref result_list 2) '( (3 . 3) 10 #f) )
      (check-equal? (list-ref result_list 3) '( (5 . 5) 10 3) )
    )

   ))

(run-tests test-range-lib)
