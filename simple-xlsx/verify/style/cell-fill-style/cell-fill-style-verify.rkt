#lang racket

(provide (contract-out
          [add-cell-fills-sheet (-> void?)]
          [set-more-cell-fills (-> void?)]
          ))

(require rackunit/text-ui rackunit)
(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_fill_file "_cell_fill.xlsx")
(define-runtime-path cell_fill_read_and_write_file "_cell_fill_read_and_write.xlsx")

(define (set-more-cell-fills)
  (with-sheet-name
   "cell fill"
   (lambda ()
     (set-cell-range-fill-style "B8-F12" "FFFF00" "solid"))))

(define (add-cell-fills-sheet)
  (add-data-sheet
   "cell fill"
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
   "cell fill"
   (lambda ()
     (set-cell-range-fill-style "B2-F6" "FF0000" "solid")
     (set-cell-range-fill-style "H2-L6" "0000FF" "gray125")
     (set-cell-range-fill-style "N2-R6" "00FF00" "darkDown")
     )))

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-fill"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_fill_file
           (lambda ()
             (add-cell-fills-sheet)))

          (read-and-write-xlsx
           cell_fill_file
           cell_fill_read_and_write_file
           (lambda ()
             (set-more-cell-fills)
             ))
          )
        (lambda ()
          ;(void)
          (delete-file cell_fill_file)
          (delete-file cell_fill_read_and_write_file)
          )))
   ))

(run-tests test-writer)
