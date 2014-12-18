#lang racket

(require xlsx-reader)

(with-input-from-excel-file
  "test1.xlsx"
  (lambda ()
    (load-sheet "Sheet1")
    (printf "~a\n" (get-cell-value "A1"))
    (printf "~a\n" (get-sheet-names))
    (printf "~a\n" (get-sheet-dimension))
    (with-row
      (lambda (row)
        (printf "~a\n" row)))))