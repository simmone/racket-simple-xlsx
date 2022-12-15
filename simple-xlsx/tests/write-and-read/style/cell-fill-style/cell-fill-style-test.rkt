#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_fill_file "cell_fill.xlsx")
(define-runtime-path cell_fill_read_and_write_file "cell_fill_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-fill"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_fill_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              (let loop-row ([rows_count 100]
                             [rows '()])
                (if (>= rows_count 1)
                    (loop-row
                     (sub1 rows_count)
                     (cons
                      (let loop-col ([cols_count 100]
                                     [cols '()])
                        (if (>= cols_count 1)
                            (loop-col
                             (sub1 cols_count)
                             (cons cols_count cols))
                            cols))
                      rows))
                    rows)))

             (with-sheet
              (lambda ()
                (set-cell-range-fill-style "B2-F6" "ff0000" "solid")
                (set-cell-range-fill-style "H2-L6" "0000ff" "gray125")
                (set-cell-range-fill-style "N2-R6" "00ff00" "darkDown")
                ))))

          (read-and-write-xlsx
           cell_fill_file
           cell_fill_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file cell_fill_file)
          (delete-file cell_fill_read_and_write_file)
          )))
   ))

(run-tests test-writer)
