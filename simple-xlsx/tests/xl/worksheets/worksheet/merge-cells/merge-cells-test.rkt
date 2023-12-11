#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../style/style.rkt")
(require "../../../../../style/styles.rkt")
(require "../../../../../style/set-styles.rkt")
(require "../../../../../lib/lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path merge_cells1_file "merge_cells1.xml")
(define-runtime-path merge_cells2_file "merge_cells2.xml")

(define test-merge-cells
  (test-suite
   "test-to-merge-cells"

   (test-case
    "test-to-merge-cells"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month1" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2"
                       '(("month1" "month2" "month3" "month1" "real") (201601 100 110 1110 6.9)))

       (with-sheet-ref
        0
        (lambda ()
          (set-merge-cell-range "A1:C2")
          (set-merge-cell-range "D3:F4")
          (set-merge-cell-range "H5-J6")

          (call-with-input-file merge_cells1_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-merge-cells))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        1
        (lambda ()
          (set-merge-cell-range "D3:F4")

          (call-with-input-file merge_cells2_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-merge-cells))
               (lambda (actual)
                 (check-lines? expected actual))))))))))

   (test-case
    "test-from-merge-cells"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((1)))
       (add-data-sheet "Sheet2" '((1)))

       (with-sheet-ref
        0
        (lambda ()

         (from-merge-cells
           (xml->hash (open-input-string (format "<worksheet>~a</worksheet>" (file->string merge_cells1_file)))))

         (check-true (hash-has-key? (SHEET-STYLE-cell_range_merge_map (*CURRENT_SHEET_STYLE*)) "A1:C2"))
         (check-true (hash-has-key? (SHEET-STYLE-cell_range_merge_map (*CURRENT_SHEET_STYLE*)) "D3:F4"))
         (check-true (hash-has-key? (SHEET-STYLE-cell_range_merge_map (*CURRENT_SHEET_STYLE*)) "H5:J6"))))

       (with-sheet-ref
        1
        (lambda ()

         (from-merge-cells
           (xml->hash (open-input-string (format "<worksheet>~a</worksheet>" (file->string merge_cells2_file)))))

         (check-true (hash-has-key? (SHEET-STYLE-cell_range_merge_map (*CURRENT_SHEET_STYLE*)) "D3:F4")))))))

   ))

(run-tests test-merge-cells)
