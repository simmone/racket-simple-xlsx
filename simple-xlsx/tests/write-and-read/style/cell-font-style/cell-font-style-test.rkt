#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_font_file "cell_font.xlsx")
(define-runtime-path cell_font_read_and_write_file "cell_font_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-cell-range-font"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_font_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                ("12" "Arial" "000000")
                ("16" "Monospace" "900000")
                ("20" "Sans" "990000")
                ))

             (with-sheet
              (lambda ()
                (set-cell-range-font-style "A1-C1" 12 "Arial" "000000")
                (set-cell-range-font-style "A2-C2" 16 "Monospace" "900000")
                (set-cell-range-font-style "A3-C3" 20 "Sans" "990000")
                (set-col-range-width "A-C" 30)
                (set-row-range-height "1-3" 50)
                ))))

          (read-and-write-xlsx
           cell_font_file
           cell_font_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file cell_font_file)
          (delete-file cell_font_read_and_write_file)
          )))

   ))

(run-tests test-writer)
