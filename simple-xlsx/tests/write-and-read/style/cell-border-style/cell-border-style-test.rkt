#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_border_file "cell_border.xlsx")
(define-runtime-path cell_border_read_and_write_file "cell_border_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-border"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_border_file
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
                (set-cell-range-border-style "B2-F6" "all" "ff0000" "thick")
                (set-cell-range-border-style "B8-F12" "left" "ff0000" "thick")
                (set-cell-range-border-style "H2-L6" "right" "ff0000" "dashed")
                (set-cell-range-border-style "H8-L12" "top" "ff0000" "double")
                (set-cell-range-border-style "N2-R6" "bottom" "ff0000" "thick")
                (set-cell-range-border-style "N8-R12" "side" "ff0000" "thick")
                ))))

          (read-and-write-xlsx
           cell_border_file
           cell_border_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file cell_border_file)
          (delete-file cell_border_read_and_write_file)
          )))
   ))

(run-tests test-writer)
