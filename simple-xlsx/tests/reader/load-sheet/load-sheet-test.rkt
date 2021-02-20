#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader/load-sheet.rkt")

(require racket/runtime-path)
(define-runtime-path sheet_file "sheet1.xml")

(define test-load-sheet-file
  (test-suite
   "test-load-sheet-file"
   
   (test-case
    "test-load-sheet-file"

    (let-values ([(dimension data_map formula_map data_type_map)
                  (load-sheet-file sheet_file)])

      (check-equal? dimension '(4 . 4))
      (check-equal? (hash-count data_map) 16)

      ))))

(run-tests test-load-sheet-file)
