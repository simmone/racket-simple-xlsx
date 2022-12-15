#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_number_file "cell_number.xlsx")
(define-runtime-path cell_number_read_and_write_file "cell_number_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-cell-range-number"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_number_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                (1 1.1 100.11)
                (2 2.2 200.12)
                (3 3.3 300.13)
                ))

             (with-sheet
              (lambda ()
                (set-cell-range-number-style "A1-C1" "0.00")
                (set-cell-range-number-style "A2-C2" "0.000")
                (set-cell-range-number-style "A3-C3" "0,000.00%")
                (set-col-range-width "A-C" 30)
                (set-row-range-height "1-3" 50)
                ))))

          (read-and-write-xlsx
           cell_number_file
           cell_number_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file cell_number_file)
          (delete-file cell_number_read_and_write_file)
          )))

   ))

(run-tests test-writer)
