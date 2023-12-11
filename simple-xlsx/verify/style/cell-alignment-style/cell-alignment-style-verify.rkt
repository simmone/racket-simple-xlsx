#lang racket

(provide (contract-out
          [add-cell-alignment-sheet (-> void?)]
          [set-more-cell-alignments (-> void?)]
          ))

(require rackunit/text-ui rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_alignment_file "_cell_alignment.xlsx")
(define-runtime-path cell_alignment_read_and_write_file "_cell_alignment_read_and_write.xlsx")

(define (add-cell-alignment-sheet)
  (add-data-sheet
   "cell alignment"
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
   "cell alignment"
   (lambda ()
     (set-cell-range-alignment-style "B2-E5" "center" "center")
     (set-cell-range-border-style "B2-E5" "side" "FF0000" "thick")

     (set-cell-range-alignment-style "G2-J5" "left" "top")
     (set-cell-range-border-style "G2-J5" "side" "FF0000" "thick")

     (set-cell-range-alignment-style "L2-O5" "right" "bottom")
     (set-cell-range-border-style "L2-O5" "side" "FF0000" "thick")

     (set-row-range-height "1-5" 30)
     )))

(define (set-more-cell-alignments)
  (with-sheet-name
   "cell alignment"
   (lambda ()
     (set-cell-range-alignment-style "B7-E10" "left" "bottom")
     (set-cell-range-border-style "B7-E10" "side" "0000FF" "thick"))))

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-cell-range-alignment"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_alignment_file
           (lambda ()
             (add-cell-alignment-sheet)
             ))

          (read-and-write-xlsx
           cell_alignment_file
           cell_alignment_read_and_write_file
           (lambda ()
             (set-more-cell-alignments)))
          )
        (lambda ()
          ;(void)
          (delete-file cell_alignment_file)
          (delete-file cell_alignment_read_and_write_file)
          )))
   ))

(run-tests test-writer)
