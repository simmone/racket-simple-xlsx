#lang racket

(provide (contract-out
          [add-col-style-sheet (-> void?)]
          [set-more-col-styles (-> void?)]
          ))

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path col_file "_col.xlsx")
(define-runtime-path col_read_and_write_file "_col_read_and_write.xlsx")

(define (add-col-style-sheet)
  (add-data-sheet
   "col style"
   '(
     ("1" "2" "3" 4 "5")
     ("1" "2" "3" 4 "5")
     ("1" "2" "3" 4 "5")
     ("1" "2" "3" 4 "5")
     ))

  (with-sheet-name
   "col style"
   (lambda ()
     (set-col-range-font-style "1" 12 "Arial" "FF0000")
     (set-col-range-font-style "2" 16 "Monospace" "00FF00")
     (set-col-range-font-style "3" 20 "Sans" "0000FF")
     (set-col-range-alignment-style "3" "left" "bottom")
     (set-col-range-number-style "D-D" "0,000.00")
     (set-col-range-fill-style "2" "FF0000" "solid")

     (set-col-range-width "A-D" 30)
     (set-row-range-height "1-3" 50)
     )))

(define (set-more-col-styles)
  (with-sheet-name
   "col style"
   (lambda ()
     (set-col-range-fill-style "5" "0000FF" "solid")
     (set-row-range-height "4" 50))))

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
             (add-col-style-sheet)))

          (read-and-write-xlsx
           col_file
           col_read_and_write_file
           (lambda ()
             (set-more-col-styles)))
          )
        (lambda ()
          ;(void)
          (delete-file col_file)
          (delete-file col_read_and_write_file)
          )))

   ))

(run-tests test-writer)
