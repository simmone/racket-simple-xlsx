#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../style/style.rkt")
(require "../../../../../style/set-styles.rkt")
(require "../../../../../lib/lib.rkt")
(require "../../../../../lib/sheet-lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
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
                       "B2")

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
                     (lists->xml (to-work-sheet))
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
