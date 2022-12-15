#lang racket

(require rackunit/text-ui rackunit)

(require"../../lib/dimension.rkt")

(define test-dimension-lib
  (test-suite
   "test-dimension-lib"

   (test-case
    "test-cell-range?"

    (check-true (cell-range? "A1-B2"))
    (check-true (cell-range? "A1:B2"))
    (check-true (cell-range? "A1"))
    (check-true (cell-range? "B2"))
    (check-false (cell-range? "ksdkf"))
    )

   (test-case
    "test-cell?"

    (check-true (cell? "A1"))
    (check-false (cell? "a"))
    (check-false (cell? "23"))
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
    )

   (test-case
    "test-range->row_col_pair"

    (check-equal? (range->row_col_pair "A1-B2") '((1 . 1) . (2 . 2)))
    (check-equal? (range->row_col_pair "A2-D7") '((2 . 1) . (7 . 4)))
    (check-equal? (range->row_col_pair "B2-A1") '((1 . 1) . (1 . 1)))
    (check-equal? (range->row_col_pair "A1") '((1 . 1) . (1 . 1)))
    (check-equal? (range->row_col_pair "A2") '((2 . 1) . (2 . 1)))
    (check-equal? (range->row_col_pair "A2-A1") '((1 . 1) . (1 . 1)))
    (check-equal? (range->row_col_pair "B1-A1") '((1 . 1) . (1 . 1)))
    )
   
   (test-case
    "test-range->capacity"

    (check-equal? (range->capacity "A1:F4") '(4 . 6))
    (check-equal? (range->capacity "B2:F4") '(3 . 5))
    )
   
   (test-case
    "test-capacity->range"
    
    (check-equal? (capacity->range '(4 . 6)) "A1:F4")

    (check-equal? (capacity->range '(4 . 6) "B1") "B1:G4")

    (check-equal? (capacity->range '(4 . 6) "C2") "C2:H5")
    )

   (test-case
    "test-col-abc->number"
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
    "test-col_number->abc"
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
    "test-to-col-range"
    (check-equal? (to-col-range "A") '(1 . 1))
    (check-equal? (to-col-range "B") '(2 . 2))
    (check-equal? (to-col-range "2") '(2 . 2))
    (check-equal? (to-col-range "C-D") '(3 . 4))
    (check-equal? (to-col-range "3-4") '(3 . 4))
    (check-equal? (to-col-range "1-26") '(1 . 26))
    (check-equal? (to-col-range "26-1") '(1 . 1))
    (check-equal? (to-col-range "A-Z") '(1 . 26))
    (check-equal? (to-col-range "Z-A") '(1 . 1))
    (check-equal? (to-col-range "A-B-C") '(1 . 1))
    (check-equal? (to-col-range "A-ksdk344") '(1 . 1))
    (check-equal? (to-col-range "sdksjdkf-%^%$#") '(1 . 1))
    (check-equal? (to-col-range "A-Z") '(1 . 26))
    (check-equal? (to-col-range "A-z") '(1 . 26))
    (check-equal? (to-col-range "A-B") '(1 . 2))
    (check-equal? (to-col-range "A") '(1 . 1))
    (check-equal? (to-col-range "7-12") '(7 . 12))
    (check-equal? (to-col-range "10") '(10 . 10))
    )

   (test-case
    "test-to-row-range"

    (check-equal? (to-row-range "1-4") '(1 . 4))

    (check-equal? (to-row-range "7-12") '(7 . 12))

    (check-equal? (to-row-range "10") '(10 . 10))

    (check-equal? (to-row-range "2-1") '(1 . 1))

    (check-equal? (to-row-range "A-B") '(1 . 1))

    (check-equal? (to-row-range "12-7") '(1 . 1)))

   (test-case
    "test-range->range_xml"

    (check-equal? (range->range_xml "C2-C10") "$C$2:$C$10")

    (check-equal? (range->range_xml "C2-Z2") "$C$2:$Z$2")

    (check-equal? (range->range_xml "AB20-AB100") "$AB$20:$AB$100"))

   (test-case
    "test-range_xml->range"

    (check-equal? (range_xml->range "$C$2:$C$10") "C2-C10")

    (check-equal? (range_xml->range "$C$2:$Z$2") "C2-Z2")

    (check-equal? (range_xml->range "$AB$20:$AB$100") "AB20-AB100"))

   (test-case
    "test-cell_range->cell_list"

    (check-equal? (cell_range->cell_list "A1-D1") '("A1" "B1" "C1" "D1"))

    (check-equal? (cell_range->cell_list "B1-D2") '("B1" "C1" "D1" "B2" "C2" "D2"))

    (check-equal? (cell_range->cell_list "A2-A1") '("A1"))

    (check-equal? (cell_range->cell_list "B1-A1") '("A1"))
    )

   (test-case
    "test-get-cell-range-side-cells"

    (let-values ([(top_cells bottom_cells left_cells right_cells) (get-cell-range-four-sides-cells "A1")])
      (check-equal? top_cells '("A1"))
      (check-equal? bottom_cells '("A1"))
      (check-equal? left_cells '("A1"))
      (check-equal? right_cells '("A1")))

    (let-values ([(top_cells bottom_cells left_cells right_cells) (get-cell-range-four-sides-cells "A1-B1")])
      (check-equal? top_cells '("A1" "B1"))
      (check-equal? bottom_cells '("A1" "B1"))
      (check-equal? left_cells '("A1"))
      (check-equal? right_cells '("B1")))

    (let-values ([(top_cells bottom_cells left_cells right_cells) (get-cell-range-four-sides-cells "A1-A2")])
      (check-equal? top_cells '("A1"))
      (check-equal? bottom_cells '("A2"))
      (check-equal? left_cells '("A1" "A2"))
      (check-equal? right_cells '("A1" "A2")))

    (let-values ([(top_cells bottom_cells left_cells right_cells) (get-cell-range-four-sides-cells "A1-A2")])
      (check-equal? top_cells '("A1"))
      (check-equal? bottom_cells '("A2"))
      (check-equal? left_cells '("A1" "A2"))
      (check-equal? right_cells '("A1" "A2")))

    (let-values ([(top_cells bottom_cells left_cells right_cells) (get-cell-range-four-sides-cells "A1-C3")])
      (check-equal? top_cells '("A1" "B1" "C1"))
      (check-equal? bottom_cells '("A3" "B3" "C3"))
      (check-equal? left_cells '("A1" "A2" "A3"))
      (check-equal? right_cells '("C1" "C2" "C3")))
    )

   ))

(run-tests test-dimension-lib)
