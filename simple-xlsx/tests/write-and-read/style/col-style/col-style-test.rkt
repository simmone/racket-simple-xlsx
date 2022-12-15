#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path col_file "col.xlsx")
(define-runtime-path col_read_and_write_file "col_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-col-range-font"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           col_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                ("12" "Arial" "000000" 100)
                ("16" "Monospace" "900000" 100000)
                ("20" "Sans" "990000" 10000.234)
                ))

             (with-sheet
              (lambda ()
                (set-col-range-font-style "1" 12 "Arial" "ff0000")
                (set-col-range-font-style "2" 16 "Monospace" "00ff00")
                (set-col-range-font-style "3" 20 "Sans" "0000ff")
                (set-col-range-alignment-style "3" "left" "bottom")
                (set-col-range-number-style "D-D" "0,000.00")
                (set-col-range-fill-style "2" "ff0000" "solid")

                (set-col-range-width "A-D" 30)
                (set-row-range-height "1-3" 50)
                ))))

          (read-and-write-xlsx
           col_file
           col_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file col_file)
          (delete-file col_read_and_write_file)
          )))

   ))

(run-tests test-writer)
