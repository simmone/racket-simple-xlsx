#lang racket

(require rackunit/text-ui
         rackunit
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt"
         "../../lib/sheet-lib.rkt")

(define test-lib
  (test-suite
   "test-lib"

   (test-case
    "test-lib"

    (with-xlsx
     (lambda ()
       (add-data-sheet "sheet1" '((1)))

       (add-data-sheet "sheet2" '((1 2)))

       (add-data-sheet "sheet3" '(
                                  (1 2 3)
                                  (4 5 6)
                                  ))

       (add-data-sheet "sheet4" '(
                                  (1 2)
                                  (3 4)
                                  (5 6)
                                  )
                       #:start_cell? "B2"
                       )

       (with-sheet-name
        "sheet1"
        (lambda ()
          (check-equal? (get-rows-count) 1)
          (check-equal? (get-cols-count) 1)

          (check-equal? (get-cell "A1") 1)
          (set-cell! "A1" 2)
          (check-equal? (get-cell "A1") 2)

          (check-equal? (get-row 1) '(2))

          (set-row! 1 '(3))
          (check-equal? (get-row 1) '(3))

          (set-row! 1 '(7 8 9))
          (check-equal? (get-row 1) '(7))

          (check-equal? (get-rows) '((7)))
          (set-rows! '((2) (4) (6)))
          (check-equal? (get-rows) '((2)))

          (check-equal? (get-col 1) '(2))
          (check-equal? (get-col "A") '(2))
          (set-col! 1 '(3))
          (set-col! "A" '(3))
          (check-equal? (get-col 1) '(3))
          (check-equal? (get-col "A") '(3))

          (check-equal? (get-cols) '((3)))
          (set-cols! '((2) (4) (6)))
          (check-equal? (get-cols) '((2)))

          (check-equal? (get-cell "A2") "")
          (check-equal? (get-cell "B1") "")))

       (with-sheet-name
        "sheet2"
        (lambda ()
          (check-equal? (get-rows-count) 1)
          (check-equal? (get-cols-count) 2)

          (check-equal? (get-rows) '((1 2)))
          (check-equal? (get-row 1) '(1 2))
          (check-equal? (get-col 1) '(1))
          (check-equal? (get-col 2) '(2))

          (check-equal? (get-cell "A1") 1)
          (check-equal? (get-cell "B1") 2)
          ))

       (with-sheet-name
        "sheet3"
        (lambda ()
          (check-equal? (get-rows-count) 2)
          (check-equal? (get-cols-count) 3)

          (check-equal? (get-rows) '(
                                     (1 2 3)
                                     (4 5 6)
                                     ))
          (check-equal? (get-row 1) '(1 2 3))
          (check-equal? (get-row 2) '(4 5 6))

          (check-equal? (get-col 1) '(1 4))
          (check-equal? (get-col 2) '(2 5))
          (check-equal? (get-col 3) '(3 6))

          (check-equal? (get-cell "A1") 1)
          (check-equal? (get-cell "C2") 6)

          (check-equal? (get-row-cells 1) '("A1" "B1" "C1"))
          (check-equal? (get-row-cells 2) '("A2" "B2" "C2"))

          (check-equal? (get-col-cells 1) '("A1" "A2"))
          (check-equal? (get-col-cells "A") '("A1" "A2"))
          (check-equal? (get-col-cells "A") '("A1" "A2"))
          (check-equal? (get-col-cells 2) '("B1" "B2"))
          (check-equal? (get-col-cells 3) '("C1" "C2"))
          (check-equal? (get-col-cells "C") '("C1" "C2"))
          ))

       (with-sheet-name
        "sheet4"
        (lambda ()
          (check-equal? (get-rows) '(
                                     (1 2)
                                     (3 4)
                                     (5 6)
                                     ))

          (check-equal? (get-cell "B2") 1)
          (check-equal? (get-cell "C4") 6)

          (check-equal? (get-row 2) '(1 2))
          (check-equal? (get-row 3) '(3 4))
          (check-equal? (get-row 4) '(5 6))

          (check-equal? (get-col 2) '(1 3 5))
          (check-equal? (get-col 3) '(2 4 6))

          (set-rows! '(
                       (1 4)
                       (2 5)
                       (3 6)
                       ))
          (check-equal? (get-rows) '(
                                     (1 4)
                                     (2 5)
                                     (3 6)
                                     ))

          (set-row! 3 '(7 8))
          (check-equal? (get-row 3) '(7 8))

          (check-equal? (get-cols) '(
                                     (1 7 3)
                                     (4 8 6)
                                     ))

          (set-cols! '((4 5 6) (1 2 3)))
          (check-equal? (get-rows) '(
                                     (4 1)
                                     (5 2)
                                     (6 3)
                                     ))

          (check-equal? (get-row 5) '("" ""))
          (check-equal? (get-col 4) '("" "" ""))
          (check-equal? (get-col "D") '("" "" ""))

          (check-equal?
           (get-range-values "B2-C4")
           '(4 1 5 2 6 3))

          ))
       )))

   (test-case
    "test-squash-shared-string-map"

    (with-xlsx
     (lambda ()
       (add-data-sheet "sheet1" '(("1" "2" "3" 4)))
       (add-data-sheet "sheet2" '(("2" "2" "2" 4)))
       (add-data-sheet "sheet3" '(("3" "2" "1" 4)))
       (add-data-sheet "sheet4" '(("1" "2" "2" 4)))
       (add-data-sheet "sheet5" '(("2" "2" "3" 4)))

       (squash-shared-strings-map)

       (check-equal? (hash-count (XLSX-shared_string->index_map (*XLSX*))) 3)
       (check-equal? (hash-count (XLSX-shared_index->string_map (*XLSX*))) 3)

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "1") 0)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 0) "1")

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "2") 1)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 1) "2")

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "3") 2)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 2) "3")
       )))

   (test-case
    "test-lib"

    (with-xlsx
     (lambda ()
       (add-data-sheet "aheet1" '((1)))

       (add-data-sheet "bheet2" '((1 2)))

       (add-data-sheet "cheet3" '(
                                  (1 2 3)
                                  (4 5 6)
                                  ))

       (add-data-sheet "dheet4" '(
                                  (1 2)
                                  (3 4)
                                  (5 6)
                                  ))

       (check-equal? (get-sheet-ref-rows-count 0) 1)
       (check-equal? (get-sheet-name-rows-count "aheet1") 1)
       (check-equal? (get-sheet-*name*-rows-count "ahee") 1)

       (check-equal? (get-sheet-ref-cols-count 0) 1)
       (check-equal? (get-sheet-name-cols-count "aheet1") 1)
       (check-equal? (get-sheet-*name*-cols-count "ahee") 1)

       (set-sheet-ref-cell! 0 "A1" 2)
       (check-equal? (get-sheet-ref-cell 0 "A1") 2)
       (set-sheet-name-cell! "aheet1" "A1" 3)
       (check-equal? (get-sheet-name-cell "aheet1" "A1") 3)
       (set-sheet-*name*-cell! "ahee" "A1" 4)
       (check-equal? (get-sheet-*name*-cell "ahee" "A1") 4)

       (set-sheet-ref-row! 0 1 '(7 8 9))
       (check-equal? (get-sheet-ref-row 0 1) '(7))
       (set-sheet-name-row! "aheet1" 1 '(8 9 10))
       (check-equal? (get-sheet-name-row "aheet1" 1) '(8))
       (set-sheet-*name*-row! "ahee" 1 '(9 10 11))
       (check-equal? (get-sheet-*name*-row "ahee" 1) '(9))

       (set-sheet-ref-rows! 2 '((4 5 6) (7 8 9)))
       (check-equal? (get-sheet-ref-rows 2) '((4 5 6) (7 8 9)))
       (set-sheet-name-rows! "cheet3" '((5 5 6) (7 8 9)))
       (check-equal? (get-sheet-name-rows "cheet3") '((5 5 6) (7 8 9)))
       (set-sheet-*name*-rows! "eet3" '((6 5 6) (7 8 9)))
       (check-equal? (get-sheet-*name*-rows "eet3") '((6 5 6) (7 8 9)))

       (set-sheet-ref-col! 3 1 '(1 2 3))
       (set-sheet-ref-col! 3 "A" '(1 2 3))
       (check-equal? (get-sheet-ref-col 3 1) '(1 2  3))
       (check-equal? (get-sheet-ref-col 3 "A") '(1 2  3))
       (set-sheet-name-col! "dheet4" 2 '(4 5 6))
       (set-sheet-name-col! "dheet4" "B" '(4 5 6))
       (check-equal? (get-sheet-name-col "dheet4" 2) '(4 5 6))
       (check-equal? (get-sheet-name-col "dheet4" "B") '(4 5 6))
       (set-sheet-*name*-col! "eet4" 1 '(7 8 9))
       (set-sheet-*name*-col! "eet4" "A" '(7 8 9))
       (check-equal? (get-sheet-*name*-col "eet4" 1) '(7 8 9))
       (check-equal? (get-sheet-*name*-col "eet4" "A") '(7 8 9))

       (check-equal? (get-sheet-ref-cols 3) '((7 8 9) (4 5 6)))

       (set-sheet-ref-cols! 3 '((1 2 3) (4 5 6)))
       (check-equal? (get-sheet-ref-cols 3) '((1 2 3) (4 5 6)))
       (set-sheet-name-cols! "dheet4" '((2 2 3) (4 5 6)))
       (check-equal? (get-sheet-name-cols "dheet4") '((2 2 3) (4 5 6)))
       (set-sheet-*name*-cols! "eet4" '((3 2 3) (4 5 6)))
       (check-equal? (get-sheet-*name*-cols "eet4") '((3 2 3) (4 5 6)))

       (check-equal? (get-sheet-ref-range-values 3 "A1-B3") '(3 4 2 5 3 6))
       (check-equal? (get-sheet-name-range-values "dheet4" "A1-B3") '(3 4 2 5 3 6))
       (check-equal? (get-sheet-*name*-range-values "eet4" "A1-B3") '(3 4 2 5 3 6))
       )))
   ))

(run-tests test-lib)
