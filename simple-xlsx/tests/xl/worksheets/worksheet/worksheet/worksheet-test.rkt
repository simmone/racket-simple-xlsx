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

(define-runtime-path worksheet_test_file "worksheet_test.xml")
(define-runtime-path worksheet_no_dimension_test_file "worksheet_no_dimension_test.xml")
(define-runtime-path worksheet_not_formated_test_file "worksheet_not_formated_test.xml")
(define-runtime-path worksheet_dimension_not_from_A1_file "worksheet_dimension_not_from_A1_test.xml")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-to-worksheet"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month1" "real") (201601 100 110 1110 6.9)))

       (squash-shared-strings-map)

       (with-sheet
        (lambda ()
          (set-col-range-width "1-4" 8)
          (set-col-range-width "E" 6)

          (set-merge-cell-range "A1:C2")
          (set-merge-cell-range "D3:F4")
          (set-merge-cell-range "H5:J6")

          (call-with-input-file worksheet_test_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-work-sheet))
               (lambda (actual)
                 (check-lines? expected actual)))))
          )))))

   (test-case
    "test-from-worksheet"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("none")))

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month1" 0)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 0 "month1")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month2" 1)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 1 "month2")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month3" 2)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 2 "month3")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "real" 3)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 3 "real")

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml-port-to-hash
                  (open-input-string (file->string worksheet_test_file))
                  '(
                    "worksheet.dimension.ref"
                    "worksheet.sheetData.row.r"
                    "worksheet.sheetData.row.spans"
                    "worksheet.sheetData.row.s"
                    "worksheet.sheetData.row.customFormat"
                    "worksheet.sheetData.row.ht"
                    "worksheet.sheetData.row.customHeight"
                    "worksheet.sheetData.row.c.r"
                    "worksheet.sheetData.row.c.s"
                    "worksheet.sheetData.row.c.t"
                    "worksheet.sheetData.row.c.v"
                    "worksheet.mergeCells.mergeCell.ref"
                    )
                  )])

            (from-work-sheet xml_hash)

            (check-equal? (DATA-SHEET-dimension (*CURRENT_SHEET*)) "A1:E2")

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
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)

            ))))))

   (test-case
    "test-from-no-dimension-worksheet"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("none")))

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month1" 0)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 0 "month1")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month2" 1)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 1 "month2")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month3" 2)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 2 "month3")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "real" 3)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 3 "real")

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml-port-to-hash
                  (open-input-string (file->string worksheet_no_dimension_test_file))
                  '(
                    "worksheet.dimension.ref"
                    "worksheet.sheetData.row.r"
                    "worksheet.sheetData.row.spans"
                    "worksheet.sheetData.row.s"
                    "worksheet.sheetData.row.customFormat"
                    "worksheet.sheetData.row.ht"
                    "worksheet.sheetData.row.customHeight"
                    "worksheet.sheetData.row.c.r"
                    "worksheet.sheetData.row.c.s"
                    "worksheet.sheetData.row.c.t"
                    "worksheet.sheetData.row.c.v"
                    "worksheet.mergeCells.mergeCell.ref"
                    ))])

            (from-work-sheet xml_hash)

            (check-equal? (DATA-SHEET-dimension (*CURRENT_SHEET*)) "A1:E2")

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
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)

            ))))))

   (test-case
    "test-from-not-formated_worksheet"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("none")))

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month1" 0)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 0 "month1")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month2" 1)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 1 "month2")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month3" 2)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 2 "month3")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "real" 3)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 3 "real")

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml-port-to-hash
                  (open-input-string (file->string worksheet_not_formated_test_file))
                  '(
                    "worksheet.dimension.ref"
                    "worksheet.sheetData.row.r"
                    "worksheet.sheetData.row.spans"
                    "worksheet.sheetData.row.s"
                    "worksheet.sheetData.row.customFormat"
                    "worksheet.sheetData.row.ht"
                    "worksheet.sheetData.row.customHeight"
                    "worksheet.sheetData.row.c.r"
                    "worksheet.sheetData.row.c.s"
                    "worksheet.sheetData.row.c.t"
                    "worksheet.sheetData.row.c.v"
                    "worksheet.mergeCells.mergeCell.ref"
                    ))])

            (from-work-sheet xml_hash)

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
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)

            ))))))

   (test-case
    "test-dimension-not-from-a1-worksheet"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("none")))

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml-port-to-hash
                  (open-input-string (file->string worksheet_dimension_not_from_A1_file))
                  '(
                    "worksheet.dimension.ref"
                    "worksheet.sheetData.row.r"
                    "worksheet.sheetData.row.spans"
                    "worksheet.sheetData.row.s"
                    "worksheet.sheetData.row.customFormat"
                    "worksheet.sheetData.row.ht"
                    "worksheet.sheetData.row.customHeight"
                    "worksheet.sheetData.row.c.r"
                    "worksheet.sheetData.row.c.s"
                    "worksheet.sheetData.row.c.t"
                    "worksheet.sheetData.row.c.v"
                    "worksheet.mergeCells.mergeCell.ref"
                    ))])

            (from-work-sheet xml_hash)

            (check-equal? (hash-count (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*))) 4)

            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1"))

            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A2"))
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B2") 100)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C2") 110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D2") 1110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)

            ))))))

   ))

(run-tests test-worksheet)
