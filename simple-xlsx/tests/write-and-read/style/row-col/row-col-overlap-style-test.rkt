#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path row_col_file "row_col.xlsx")
(define-runtime-path row_col_read_and_write_file "row_col_read_and_write.xlsx")
(define-runtime-path col_row_file "col_row.xlsx")
(define-runtime-path col_row_read_and_write_file "col_row_read_and_write.xlsx")

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
             (add-data-sheet
              "Sheet1"
              '(
                ("1" "2" "3" 10000)
                (4 5 6 20000)
                ("7" "8" "9" 30000)
                 ))

             (with-sheet
              (lambda ()
                (set-row-range-font-style "2" 20 "Arial" "ff0000")
                (set-row-range-alignment-style "2" "left" "top")
                (set-row-range-number-style "2" "0.00%")
                (set-row-range-fill-style "2" "0000ff" "gray125")

                (set-col-range-font-style "4" 20 "Sans" "0000ff")
                (set-col-range-alignment-style "4" "right" "bottom")
                (set-col-range-number-style "4" "0,000.00")
                (set-col-range-fill-style "4" "ff0000" "gray125")

                (set-col-range-width "A-D" 30)
                (set-row-range-height "1-3" 50)
                ))))

          (read-and-write-xlsx
           row_col_file
           row_col_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file row_col_file)
          (delete-file row_col_read_and_write_file)
          )))

   (test-case
    "test-col_row-range-font"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           col_row_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                ("1" "2" "3" 10000)
                (4 5 6 20000)
                ("7" "8" "9" 30000)
                ))

             (with-sheet
              (lambda ()
                (set-col-range-font-style "4" 20 "Sans" "0000ff")
                (set-col-range-alignment-style "4" "right" "bottom")
                (set-col-range-number-style "4" "0,000.00")
                (set-col-range-fill-style "4" "ff0000" "gray125")

                (set-row-range-font-style "2" 20 "Arial" "ff0000")
                (set-row-range-alignment-style "2" "left" "top")
                (set-row-range-number-style "2" "0.00%")
                (set-row-range-fill-style "2" "0000ff" "gray125")

                (set-col-range-width "A-D" 30)
                (set-row-range-height "1-3" 50)
                ))))

          (read-and-write-xlsx
           col_row_file
           col_row_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file col_row_file)
          (delete-file col_row_read_and_write_file)
          )))

   ))

(run-tests test-writer)
