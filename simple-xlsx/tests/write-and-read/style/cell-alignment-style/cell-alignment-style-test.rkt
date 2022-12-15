#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_alignment_file "cell_alignment.xlsx")
(define-runtime-path cell_alignment_read_and_write_file "cell_alignment_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-cell-range-alignment"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_alignment_file
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
                (set-cell-range-alignment-style "A1-E5" "center" "center")
                (set-cell-range-border-style "A1-E5" "side" "ff0000" "thick")

                (set-cell-range-alignment-style "G1-K5" "left" "top")
                (set-cell-range-border-style "G1-K5" "side" "ff0000" "thick")

                (set-cell-range-alignment-style "M1-Q5" "right" "bottom")
                (set-cell-range-border-style "M1-Q5" "side" "ff0000" "thick")

                (set-row-range-height "1-5" 30)

                ))))

          (read-and-write-xlsx
           cell_alignment_file
           cell_alignment_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file cell_alignment_file)
          (delete-file cell_alignment_read_and_write_file)
          )))

   ))

(run-tests test-writer)
