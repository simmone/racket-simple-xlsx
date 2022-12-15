#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path row_file "row.xlsx")
(define-runtime-path row_read_and_write_file "row_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-row-range-font"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           row_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                ("12" "Arial" "000000")
                ("16" "Monospace" "900000")
                ("20" "Sans" "990000")
                (100 1000 10000000)
                ))

             (with-sheet
              (lambda ()
                (set-row-range-font-style "1" 12 "Arial" "ff0000")
                (set-row-range-font-style "2" 16 "Monospace" "00ff00")
                (set-row-range-font-style "3" 20 "Sans" "0000ff")
                (set-row-range-alignment-style "1" "center" "center")
                (set-row-range-number-style "4" "0,000.00")
                (set-row-range-fill-style "2" "ff0000" "solid")

                (set-col-range-width "A-C" 30)
                (set-row-range-height "1-4" 50)
                ))))

          (read-and-write-xlsx
           row_file
           row_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file row_file)
          (delete-file row_read_and_write_file)
          )))

   ))

(run-tests test-writer)
