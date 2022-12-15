#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path freeze_file "freeze.xlsx")
(define-runtime-path freeze_read_and_write_file "freeze_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-freeze"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           "freeze.xlsx"
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
                (set-freeze-row-col-range 2 2)))))

          (read-and-write-xlsx
           freeze_file
           freeze_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file freeze_file)
          (delete-file freeze_read_and_write_file)
          )))
   ))

(run-tests test-writer)
