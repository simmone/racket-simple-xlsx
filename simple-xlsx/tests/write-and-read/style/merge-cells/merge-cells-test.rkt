#lang racket

(require rackunit/text-ui rackunit)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path merge_cells_file "merge_cells.xlsx")
(define-runtime-path merge_cells_read_and_write_file "merge_cells_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-merge_cells"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           "merge_cells.xlsx"
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                ("Topic1" "" "" "" "" "")
                ("Topic2" "" "" "Topic2" "" "")
                ("Topic3" "" "Topic3" "" "Topic3" "")
                (1 2 3 4 5 6)))

             (with-sheet
              (lambda ()
                (set-merge-cell-range "A1-F1")
                (set-merge-cell-range "A2-C2")
                (set-merge-cell-range "D2-F2")
                (set-merge-cell-range "A3-B3")
                (set-merge-cell-range "C3-D3")
                (set-merge-cell-range "E3-F3")
                (set-row-range-alignment-style "1-4" "center" "center")
                ))))

          (read-and-write-xlsx
           merge_cells_file
           merge_cells_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file merge_cells_file)
          (delete-file merge_cells_read_and_write_file)
          )))
   ))

(run-tests test-writer)
