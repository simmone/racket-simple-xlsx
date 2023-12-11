#lang racket

(provide (contract-out
          [add-merge-cells-style-sheet (-> void?)]
          [set-more-merge-cells-styles (-> void?)]
          ))

(require rackunit/text-ui rackunit)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path merge_cells_file "_merge_cells.xlsx")
(define-runtime-path merge_cells_read_and_write_file "_merge_cells_read_and_write.xlsx")

(define (add-merge-cells-style-sheet)
  (add-data-sheet
   "merge cells"
   '(
     ("Topic1" "" "" "" "" "")
     ("Topic2" "" "" "Topic2" "" "")
     ("Topic3" "" "Topic3" "" "Topic3" "")
     (1 2 3 4 5 6)))

  (with-sheet-name
   "merge cells"
   (lambda ()
     (set-merge-cell-range "A1-F1")
     (set-merge-cell-range "A2-C2")
     (set-merge-cell-range "D2-F2")
     (set-merge-cell-range "A3-B3")
     (set-merge-cell-range "C3-D3")
     (set-merge-cell-range "E3-F3")
     (set-row-range-alignment-style "1-4" "center" "center")
     )))

(define (set-more-merge-cells-styles)
  (with-sheet-name
   "merge cells"
   (lambda ()
     (set-merge-cell-range "A4-F4"))))

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-merge_cells"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           merge_cells_file
           (lambda ()
             (add-merge-cells-style-sheet)))

          (read-and-write-xlsx
           merge_cells_file
           merge_cells_read_and_write_file
           (lambda ()
             (set-more-merge-cells-styles)))
          )
        (lambda ()
          ;(void)
          (delete-file merge_cells_file)
          (delete-file merge_cells_read_and_write_file)
          )))
   ))

(run-tests test-writer)
