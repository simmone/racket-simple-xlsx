#lang racket

(provide (contract-out
          [add-cell-border-sheet (-> void?)]
          [set-more-cell-borders (-> void?)]
          ))

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path cell_border_file "_cell_border.xlsx")
(define-runtime-path cell_border_read_and_write_file "_cell_border_read_and_write.xlsx")

(define (add-cell-border-sheet)
  (add-data-sheet
   "cell border"
   (let loop-row ([rows_count 100]
                  [rows '()])
     (if (>= rows_count 1)
         (loop-row
          (sub1 rows_count)
          (cons
           (let loop-col ([cols_count 100]
                          [cols '()])
             (if (>= cols_count 1)
                 (loop-col
                  (sub1 cols_count)
                  (cons cols_count cols))
                 cols))
           rows))
         rows)))

  (with-sheet-name
   "cell border"
   (lambda ()
     (set-cell-range-border-style "B2-F6" "all" "FF0000" "thick")
     (set-cell-range-border-style "B8-F12" "left" "FF0000" "thick")
     (set-cell-range-border-style "H2-L6" "right" "FF0000" "dashed")
     (set-cell-range-border-style "H8-L12" "top" "FF0000" "double")
     (set-cell-range-border-style "N2-R6" "bottom" "FF0000" "thick")
     (set-cell-range-border-style "N8-R12" "side" "FF0000" "thick")
     )))

(define (set-more-cell-borders)
  (with-sheet-name
   "cell border"
   (lambda ()
     (set-cell-range-border-style "B14-F18" "all" "0000FF" "thin"))))

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-border"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_border_file
           (lambda ()
             (add-cell-border-sheet)))

          (read-and-write-xlsx
           cell_border_file
           cell_border_read_and_write_file
           (lambda ()
             (set-more-cell-borders)))
          )
        (lambda ()
          ;(void)
          (delete-file cell_border_file)
          (delete-file cell_border_read_and_write_file)
          )))
   ))

(run-tests test-writer)
