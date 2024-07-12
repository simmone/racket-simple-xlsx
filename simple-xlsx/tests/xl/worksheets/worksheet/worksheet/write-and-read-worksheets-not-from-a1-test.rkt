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

(define test-worksheets
  (test-suite
   "test-worksheets"

   (test-case
    "test-write-worksheets"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month1" "real") (201601 100 110 1110 6.9))
                       #:start_cell? "B2")

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

             (read-worksheets (apply build-path (drop-right (explode-path sheet1_file) 1)))

             (with-sheet-ref
              0
              (lambda ()
                (check-equal? (hash-count (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))) 10)

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B2") "month1")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C2") "month2")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D2") "month3")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") "month1")
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F2") "real")

                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B3") 201601)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C3") 100)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D3") 110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E3") 1110)
                (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F3") 6.9)))
             )
           (lambda ()
;;             (void)
             (when (file-exists? sheet1_file) (delete-file sheet1_file))
             )))))
   ))

(run-tests test-worksheets)
