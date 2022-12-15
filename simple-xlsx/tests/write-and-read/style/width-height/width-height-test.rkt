#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path width_height_file "width_height.xlsx")
(define-runtime-path width_height_read_and_write_file "width_height_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-width_height"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           "width_height.xlsx"
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))

             (with-sheet
              (lambda ()
                (set-col-range-width "A-C" 20)
                (set-row-range-height "1-2" 40)))
             ))

          (read-and-write-xlsx
           width_height_file
           width_height_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file width_height_file)
          (delete-file width_height_read_and_write_file)
          )))
   ))

(run-tests test-writer)
