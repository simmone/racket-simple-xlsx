#lang racket

(provide (contract-out
          [add-cell-number-sheet (-> void?)]
          [set-more-cell-numbers (-> void?)]
          ))

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path cell_number_file "_cell_number.xlsx")
(define-runtime-path cell_number_read_and_write_file "_cell_number_read_and_write.xlsx")

(define (add-cell-number-sheet)
  (add-data-sheet
   "cell number"
   '(
     (1 1.1 100.11)
     (2 2.2 200.12)
     (3 3.3 300.13)
     (4 0.90125 0.90124)
     ))

  (with-sheet-name
   "cell number"
   (lambda ()
     (set-cell-range-number-style "A1-C1" "0.00")
     (set-cell-range-number-style "A2-C2" "0.000")
     (set-cell-range-number-style "A3-C3" "0,000.00%")
     (set-col-range-width "A-C" 30)
     (set-row-range-height "1-4" 50)
     )))

(define (set-more-cell-numbers)
  (with-sheet-name
   "cell number"
   (lambda ()
     (set-cell-range-number-style "A4-C4" "0.00%"))))

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
             (add-cell-number-sheet)
             ))

          (read-and-write-xlsx
           cell_number_file
           cell_number_read_and_write_file
           (lambda ()
             (set-more-cell-numbers)))
          )
        (lambda ()
          ;(void)
          (delete-file cell_number_file)
          (delete-file cell_number_read_and_write_file)
          )))

   ))

(run-tests test-writer)
