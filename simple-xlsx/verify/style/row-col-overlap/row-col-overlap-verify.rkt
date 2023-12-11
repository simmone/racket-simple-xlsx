#lang racket

(provide (contract-out
          [add-row-col-style-sheet (-> void?)]
          [set-more-row-col-styles (-> void?)]
          [add-col-row-style-sheet (-> void?)]
          [set-more-col-row-styles (-> void?)]
          ))

(require rackunit/text-ui rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path row_col_file "_row_col.xlsx")
(define-runtime-path row_col_read_and_write_file "_row_col_read_and_write.xlsx")
(define-runtime-path col_row_file "_col_row.xlsx")
(define-runtime-path col_row_read_and_write_file "_col_row_read_and_write.xlsx")

(define (add-row-col-style-sheet)
  (add-data-sheet
   "row col overlap"
   '(
     ("1" "2" "3" 10000)
     (4 5 6 20000)
     ("7" "8" "9" 30000)
     ))

  (with-sheet-name
   "row col overlap"
   (lambda ()
     (set-row-range-font-style "2" 20 "Arial" "FF0000")
     (set-row-range-alignment-style "2" "left" "top")
     (set-row-range-number-style "2" "0.00%")
     (set-row-range-fill-style "2" "0000FF" "gray125")

     (set-col-range-font-style "4" 20 "Sans" "0000FF")
     (set-col-range-alignment-style "4" "right" "bottom")
     (set-col-range-number-style "4" "0,000.00")
     (set-col-range-fill-style "4" "FF0000" "gray125")

     (set-col-range-width "A-D" 30)
     (set-row-range-height "1-3" 50)
     )))

(define (set-more-row-col-styles)
  (with-sheet-name
   "row col overlap"
   (lambda ()
     (set-col-range-fill-style "3" "0000FF" "solid"))))

(define (add-col-row-style-sheet)
  (add-data-sheet
   "col row overlap"
   '(
     ("1" "2" "3" 10000)
     (4 5 6 20000)
     ("7" "8" "9" 30000)
     ))

  (with-sheet-name
   "col row overlap"
   (lambda ()
     (set-col-range-font-style "4" 20 "Sans" "0000FF")
     (set-col-range-alignment-style "4" "right" "bottom")
     (set-col-range-number-style "4" "0,000.00")
     (set-col-range-fill-style "4" "FF0000" "gray125")

     (set-row-range-font-style "2" 20 "Arial" "FF0000")
     (set-row-range-alignment-style "2" "left" "top")
     (set-row-range-number-style "2" "0.00%")
     (set-row-range-fill-style "2" "0000FF" "gray125")

     (set-col-range-width "A-D" 30)
     (set-row-range-height "1-3" 50)
     )))

(define (set-more-col-row-styles)
  (with-sheet-name
   "col row overlap"
   (lambda ()
     (set-row-range-fill-style "3" "000FF0" "solid"))))

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-row_col-range-font"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           row_col_file
           (lambda ()
             (add-row-col-style-sheet)))

          (read-and-write-xlsx
           row_col_file
           row_col_read_and_write_file
           (lambda ()
             (set-more-row-col-styles))))
        (lambda ()
          (void)
          ;(delete-file row_col_file)
          ;(delete-file row_col_read_and_write_file)
          )))

   (test-case
    "test-col_row-range-font"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           col_row_file
           (lambda ()
             (add-col-row-style-sheet)))

          (read-and-write-xlsx
           col_row_file
           col_row_read_and_write_file
           (lambda ()
             (set-more-col-row-styles)))
          )
        (lambda ()
          ;(void)
          (delete-file col_row_file)
          (delete-file col_row_read_and_write_file)
          (delete-file row_col_file)
          (delete-file row_col_read_and_write_file)
          )))

   ))

(run-tests test-writer)
