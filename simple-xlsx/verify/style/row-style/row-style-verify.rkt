#lang racket

(provide (contract-out
          [add-row-style-sheet (-> void?)]
          [set-more-row-styles (-> void?)]
          ))

(require rackunit/text-ui rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path row_file "_row.xlsx")
(define-runtime-path row_read_and_write_file "_row_read_and_write.xlsx")

(define (add-row-style-sheet)
  (add-data-sheet
   "row style"
   '(
     ("fontSize: 12" "family: Arial" "color: FF0000" "center:center")
     ("fontSize: 16" "family: Monospace" "color: 00FF00" "left:center")
     ("fontSize: 24" "family: Sans" "color: 0000FF" "left:center")
     ("fontSize: 32" 1000 10000000 "left:center")
     ("fontSize: 48" 1000 10000000 "right:bottom")
     ))

  (with-sheet-name
   "row style"
   (lambda ()
     (set-row-range-font-style "1" 12 "Arial" "FF0000")
     (set-row-range-font-style "2" 16 "Monospace" "00FF00")
     (set-row-range-font-style "3" 24 "Sans" "0000FF")
     (set-row-range-font-style "4" 32 "Arial" "F0F0F0")

     (set-row-range-alignment-style "1" "center" "center")
     (set-row-range-number-style "4" "0,000.00")
     (set-row-range-fill-style "2" "FF0000" "solid")

     (set-col-range-width "A-D" 50)
     (set-row-range-height "1-3" 50)
     (set-row-range-height "4" 80)
     )))

(define (set-more-row-styles)
  (with-sheet-name
   "row style"
   (lambda ()
     (set-row-range-font-style "5" 48 "Arial" "0F0F0F")
     (set-row-range-alignment-style "5" "right" "bottom")
     (set-row-range-height "5" 100)
     (set-col-range-width "A-D" 50)
     )))

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
             (add-row-style-sheet)))

          (read-and-write-xlsx
           row_file
           row_read_and_write_file
           (lambda ()
             (set-more-row-styles)))
          )
        (lambda ()
          ;(void)
          (delete-file row_file)
          (delete-file row_read_and_write_file)
          )))
   ))

(run-tests test-writer)
