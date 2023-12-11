#lang racket

(provide (contract-out
          [add-cell-font-sheet (-> void?)]
          [set-more-cell-fonts (-> void?)]
          ))

(require rackunit/text-ui rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_font_file "_cell_font.xlsx")
(define-runtime-path cell_font_read_and_write_file "_cell_font_read_and_write.xlsx")

(define (add-cell-font-sheet)
  (add-data-sheet
   "cell font"
   '(
     ("12" "Arial Regular" "000000")
     ("16" "Monoco" "900000")
     ("20" "Rockwell Regular" "990000")
     ("24" "Apple Chancery" "FF0000")
     ))

  (with-sheet-name
   "cell font"
   (lambda ()
     (set-cell-range-font-style "A1-C1" 12 "Arial Regular" "000000")
     (set-cell-range-font-style "A2-C2" 16 "Monoco" "900000")
     (set-cell-range-font-style "A3-C3" 20 "Rockwell Regular" "990000")
     (set-col-range-width "A-C" 30)
     (set-row-range-height "1-3" 50)
     )))

(define (set-more-cell-fonts)
  (with-sheet-name
   "cell font"
   (lambda ()
     (set-cell-range-font-style "A4-C4" 24 "Apple Chancery" "FF0000")
     (set-row-range-height "4-4" 50))))

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
             (add-cell-font-sheet)))

          (read-and-write-xlsx
           cell_font_file
           cell_font_read_and_write_file
           (lambda ()
             (set-more-cell-fonts)))
          )
        (lambda ()
          ;(void)
          (delete-file cell_font_file)
          (delete-file cell_font_read_and_write_file)
          )))
   ))

(run-tests test-writer)
