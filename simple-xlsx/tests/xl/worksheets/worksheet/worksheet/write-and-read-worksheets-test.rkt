#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../style/style.rkt"
         "../../../../../style/set-styles.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../lib/sheet-lib.rkt"
         "../../../../../xl/worksheets/worksheet.rkt"
         racket/runtime-path)

(define-runtime-path sheet1_file "sheet1.xml")
(define-runtime-path sheet2_file "sheet2.xml")
(define-runtime-path sheet3_file "sheet3.xml")

(define test-worksheets
  (test-suite
   "test-worksheets"

   (test-case
    "test-write-worksheets"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month1" "real") (201601 100 110 1110 6.9)))

       (with-sheet-ref
        0
        (lambda ()
          (set-merge-cell-range "A1:C2")
          (set-merge-cell-range "D3:F4")
          (set-merge-cell-range "H5-J6")

          (set-col-range-width "1-4" 8)
          (set-col-range-width "5-5" 6)))

       (add-data-sheet "Sheet2"
                       '(("month2" "month2" "month3" "month1" "real") (201601 100 110 1110 6.9)))
       (with-sheet-ref
        1
        (lambda ()
          (set-merge-cell-range "A1:C2")
          (set-merge-cell-range "D3:F4")
          (set-merge-cell-range "H5-J6")

          (set-col-range-width "1-4" 8)
          (set-col-range-width "5-5" 6)))

       (add-chart-sheet "Chart" 'LINE "Chart1" '())

       (add-data-sheet "Sheet3"
                       '(("month3" "month2" "month3" "month1" "real") (201601 100 110 1110 6.9)))

       (with-sheet-ref
        3
        (lambda ()
          (set-merge-cell-range "A1:C2")
          (set-merge-cell-range "D3:F4")
          (set-merge-cell-range "H5-J6")

          (set-col-range-width "1-4" 8)
          (set-col-range-width "5-5" 6)))

       (squash-shared-strings-map)

       (dynamic-wind
           (lambda ()
             (write-worksheets (apply build-path (drop-right (explode-path sheet1_file) 1))))
           (lambda ()
             (call-with-input-file sheet1_file
               (lambda (expected)
                 (with-sheet-ref
                  0
                  (lambda ()
                    (call-with-input-string
                     (lists-to-xml (to-work-sheet))
                     (lambda (actual)
                       (check-lines? expected actual)))))))

             (call-with-input-file sheet2_file
               (lambda (expected)
                 (with-sheet-ref
                  1
                  (lambda ()
                    (call-with-input-string
                     (lists-to-xml (to-work-sheet))
                     (lambda (actual)
                       (check-lines? expected actual)))))))

             (call-with-input-file sheet3_file
               (lambda (expected)
                 (with-sheet-ref
                  3
                  (lambda ()
                    (call-with-input-string
                     (lists-to-xml (to-work-sheet))
                     (lambda (actual)
                       (check-lines? expected actual)))))))

             (read-worksheets (apply build-path (drop-right (explode-path sheet1_file) 1)))

             (with-sheet-ref
              0
              (lambda ()
                (check-equal? (hash-count (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))) 10)

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "month1")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "month2")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "month3")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "month1")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "real")

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A2") 201601)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B2") 100)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C2") 110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D2") 1110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)))

             (with-sheet-ref
              1
              (lambda ()
                (check-equal? (hash-count (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))) 10)

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "month2")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "month2")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "month3")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "month1")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "real")

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A2") 201601)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B2") 100)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C2") 110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D2") 1110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)))

             (with-sheet-ref
              3
              (lambda ()
                (check-equal? (hash-count (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))) 10)

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "month3")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "month2")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "month3")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "month1")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "real")

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A2") 201601)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B2") 100)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C2") 110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D2") 1110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)))

             )
           (lambda ()
             (when (file-exists? sheet1_file) (delete-file sheet1_file))
             (when (file-exists? sheet2_file) (delete-file sheet2_file))
             (when (file-exists? sheet3_file) (delete-file sheet3_file))
             )))))
   ))

(run-tests test-worksheets)
