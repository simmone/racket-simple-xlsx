#lang racket

(provide (contract-out
          [add-width-height-style-sheet (-> void?)]
          [set-more-width-height-styles (-> void?)]
          ))

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path width_height_file "_width_height.xlsx")
(define-runtime-path width_height_read_and_write_file "_width_height_read_and_write.xlsx")

(define (add-width-height-style-sheet)
  (add-data-sheet
   "width height"
   '(
     ("A1" "B1" "C1" "D1")
     ("A2" "B2" "C2" "D2")
     ("A3" "B3" "C3" "D3")
     ))
  
  (with-sheet-name
   "width height"
   (lambda ()
     (set-col-range-width "A-C" 20)
     (set-row-range-height "1-2" 40)))
  )

(define (set-more-width-height-styles)
  (with-sheet-name
   "width height"
   (lambda ()
     (set-col-range-width "D" 40)
     (set-row-range-height "3" 60))))

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-width_height"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           width_height_file
           (lambda ()
             (add-width-height-style-sheet)))

          (read-and-write-xlsx
           width_height_file
           width_height_read_and_write_file
           (lambda ()
             (set-more-width-height-styles)))
          )
        (lambda ()
          ;(void)
          (delete-file width_height_file)
          (delete-file width_height_read_and_write_file)
          )))
   ))

(run-tests test-writer)
