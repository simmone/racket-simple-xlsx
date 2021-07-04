#lang racket

(require rackunit/text-ui)

(require rackunit "../../lib/dimension.rkt")

(define test-dimension-lib
  (test-suite
   "test-dimension-lib"
   
   (test-case
    "test-get-dimension"
    
    (check-equal? (get-dimension '((1 2) (3 4))) '(2 . 2))
    (check-equal? (get-dimension '((1 2 3) (3 4 6))) '(2 . 3)))
   
   (test-case
    "test-dimension->pair"
    
    (check-equal? (dimension->pair "A1:F4") '(4 . 6))
    )
   
   (test-case
    "test-row_col->cell"
    
    (check-equal? (row_col->cell 1 1) "A1")
    (check-equal? (row_col->cell 2 1) "A2")
    (check-equal? (row_col->cell 2 5) "E2"))

   (test-case
    "test-cell->row_col"
    
    (let ([rowcol (cell->row_col "A1")])
      (check-equal? (car rowcol) 1)
      (check-equal? (cdr rowcol) 1))

    (let ([rowcol (cell->row_col "A2")])
      (check-equal? (car rowcol) 2)
      (check-equal? (cdr rowcol) 1))

    (let ([rowcol (cell->row_col "E2")])
      (check-equal? (car rowcol) 2)
      (check-equal? (cdr rowcol) 5))

    (let ([rowcol (cell->row_col "C10")])
      (check-equal? (car rowcol) 10)
      (check-equal? (cdr rowcol) 3))

    (let ([rowcol (cell->row_col "AB23")])
      (check-equal? (car rowcol) 23)
      (check-equal? (cdr rowcol) 28))

    (let ([rowcol (cell->row_col "23")])
      (check-equal? (car rowcol) 0)
      (check-equal? (cdr rowcol) 0))

    (let ([rowcol (cell->row_col "A")])
      (check-equal? (car rowcol) 0)
      (check-equal? (cdr rowcol) 0))
    )


   (test-case 
    "test-AZ-NUMBER"
    (check-equal? (col_abc->number "A") 1)
    (check-equal? (col_abc->number "B") 2)
    (check-equal? (col_abc->number "Z") 26)
    (check-equal? (col_abc->number "AA") 27)
    (check-equal? (col_abc->number "AB") 28)
    (check-equal? (col_abc->number "AZ") 52)
    (check-equal? (col_abc->number "BA") 53)
    (check-equal? (col_abc->number "YZ") 676)
    (check-equal? (col_abc->number "ZA") 677)
    (check-equal? (col_abc->number "ZZ") 702)
    (check-equal? (col_abc->number "AAA") 703)
   )

   (test-case 
    "test-NUMBER-AZ"
    (check-equal? (col_number->abc 1) "A")
    (check-equal? (col_number->abc 26) "Z")
    (check-equal? (col_number->abc 27) "AA")
    (check-equal? (col_number->abc 28) "AB")
    (check-equal? (col_number->abc 29) "AC")
    (check-equal? (col_number->abc 51) "AY")
    (check-equal? (col_number->abc 52) "AZ")
    (check-equal? (col_number->abc 53) "BA")
    (check-equal? (col_number->abc 676) "YZ")
    (check-equal? (col_number->abc 677) "ZA")
    (check-equal? (col_number->abc 702) "ZZ")
    (check-equal? (col_number->abc 703) "AAA")
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

    (check-equal? (check-col-range "7-12") "7-12")

    (check-equal? (check-col-range "10") "10-10")
    
    (check-exn exn:fail? (lambda () (check-col-range "B-A")))

    (check-exn exn:fail? (lambda () (check-col-range "A1-A")))
    )

   (test-case
    "test-check-row-range"
    
    (check-equal? (check-row-range "1-4") '(1 . 4))

    (check-equal? (check-row-range "7-12") '(7 . 12))

    (check-equal? (check-row-range "10") '(10 . 10))

    (check-exn exn:fail? (lambda () (check-row-range "2-1")))

    (check-exn exn:fail? (lambda () (check-row-range "A-B")))

    (check-exn exn:fail? (lambda () (check-row-range "12-7")))
    )

   (test-case
    "test-check-cell-range"
    
    (check-cell-range "A1-B2")

    (check-cell-range "G7")

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

    ))


   (test-case
    "test-cross-cell-style"
    
    (let ([row_map (make-hash)]
          [col_map (make-hash)]
          [row_style_map (make-hash)]
          [col_style_map (make-hash)])

      (hash-set! row_style_map 'a 1)
      (hash-set! row_style_map 'b 2)

      (hash-set! col_style_map 'a 2)
      (hash-set! col_style_map 'b 1)
      (hash-set! col_style_map 'c 3)

      (hash-set! row_map 3 row_style_map)
      (hash-set! row_map 4 row_style_map)

      (hash-set! col_map 1 col_style_map)
      (hash-set! col_map 3 col_style_map)
      
      (define row_end_map (cross-cell-style row_map col_map 'row))
      (check-equal? (hash-count row_end_map) 4)
      (check-true (hash-has-key? row_end_map "A3"))
      (check-true (hash-has-key? row_end_map "A4"))
      (check-true (hash-has-key? row_end_map "C3"))
      (check-true (hash-has-key? row_end_map "C4"))
      
      (let ([cell_style_map (hash-ref row_end_map "C3")])
        (check-equal? (hash-count cell_style_map) 3)
        (check-equal? (hash-ref cell_style_map 'a) 1)
        (check-equal? (hash-ref cell_style_map 'b) 2)
        (check-equal? (hash-ref cell_style_map 'c) 3))

      (define col_end_map (cross-cell-style row_map col_map 'col))
      (check-equal? (hash-count col_end_map) 4)
      (check-true (hash-has-key? col_end_map "A3"))
      (check-true (hash-has-key? col_end_map "A4"))
      (check-true (hash-has-key? col_end_map "C3"))
      (check-true (hash-has-key? col_end_map "C4"))
      
      (let ([cell_style_map (hash-ref col_end_map "C3")])
        (check-equal? (hash-count cell_style_map) 3)
        (check-equal? (hash-ref cell_style_map 'a) 2)
        (check-equal? (hash-ref cell_style_map 'b) 1)
        (check-equal? (hash-ref cell_style_map 'c) 3))

      )
    )

   (test-case
    "test-expand-row-style-to-cell"
    
    (let ([cells_map (make-hash)]
          [cell_style_map (make-hash)]
          [rows_map (make-hash)]
          [row_style_map (make-hash)])

      (hash-set! cell_style_map 'a 1)
      (hash-set! cell_style_map 'b 2)
      (hash-set! cells_map "A5" cell_style_map)

      (hash-set! row_style_map 'b 3)

      (hash-set! rows_map 1 row_style_map)
      (expand-row-style-to-cell rows_map cells_map)
      (check-equal? (hash-ref (hash-ref cells_map "A5") 'b) 2)

      (hash-set! rows_map 5 row_style_map)
      (expand-row-style-to-cell rows_map cells_map)
      (check-equal? (hash-ref (hash-ref cells_map "A5") 'b) 3)
      ))

   (test-case
    "test-expand-col-style-to-cell"
    
    (let ([cells_map (make-hash)]
          [cell_style_map (make-hash)]
          [cols_map (make-hash)]
          [col_style_map (make-hash)])

      (hash-set! cell_style_map 'a 1)
      (hash-set! cell_style_map 'b 2)
      (hash-set! cells_map "A5" cell_style_map)

      (hash-set! col_style_map 'b 3)

      (hash-set! cols_map 2 col_style_map)
      (expand-col-style-to-cell cols_map cells_map)
      (check-equal? (hash-ref (hash-ref cells_map "A5") 'b) 2)

      (hash-set! cols_map 1 col_style_map)
      (expand-col-style-to-cell cols_map cells_map)
      (check-equal? (hash-ref (hash-ref cells_map "A5") 'b) 3)
      ))

   ))

(run-tests test-dimension-lib)
