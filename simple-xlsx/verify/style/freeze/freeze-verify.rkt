#lang racket

(provide (contract-out
          [add-freeze-style-sheet (-> void?)]
          [set-more-freeze-styles (-> void?)]
          ))

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path freeze_file "_freeze.xlsx")
(define-runtime-path freeze_read_and_write_file "_freeze_read_and_write.xlsx")

(define (add-freeze-style-sheet)
  (add-data-sheet
   "freeze"
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
   "freeze"
   (lambda ()
     (set-freeze-row-col-range 2 2))))

(define (set-more-freeze-styles)
  (with-sheet-name
   "freeze"
   (lambda ()
     (set-freeze-row-col-range 3 3))))

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-freeze"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           freeze_file
           (lambda ()
             (add-freeze-style-sheet)))

          (read-and-write-xlsx
           freeze_file
           freeze_read_and_write_file
           (lambda ()
             (set-more-freeze-styles)))
          )
        (lambda ()
          ;(void)
          (delete-file freeze_file)
          (delete-file freeze_read_and_write_file)
          )))
   ))

(run-tests test-writer)
