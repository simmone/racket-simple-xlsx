#lang racket

(require simple-xml)

(require racket/date)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../style/style.rkt")
(require "../../../../../style/border-style.rkt")
(require "../../../../../style/styles.rkt")
(require "../../../../../style/assemble-styles.rkt")
(require "../../../../../style/set-styles.rkt")
(require "../../../../../lib/lib.rkt")
(require "../../../../../lib/sheet-lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path row1_file "row1.xml")
(define-runtime-path row2_file "row2.xml")
(define-runtime-path row3_file "row3.xml")
(define-runtime-path row1_not_from_a1_file "row1_not_from_a1.xml")
(define-runtime-path row_with_oversize_dimension_file "row_with_oversize_dimension.xml")

(define test-row
  (test-suite
   "test-row"

   (test-case
    "test-to-row"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       (list
                        (list "month1" "month2" "month3" "month1" "real")
                        (list 201601 100 110 1110 6.9)
                        (list
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                         (seconds->date (find-seconds 0 0 0 17 9 2018 #f)))
                        ))

       (with-sheet
        (lambda ()
          (set-row-range-height "1-3" 5)
          (set-row-range-font-style "1" 10 "Arial" "0000FF")))

       (squash-shared-strings-map)

       (strip-styles)

       (with-sheet
        (lambda ()
          (call-with-input-file row1_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-row 1 1 5))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file row2_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-row 2 1 5))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file row3_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-row 3 1 5))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file row_with_oversize_dimension_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-row 1 1 10))
               (lambda (actual)
                 (check-lines? expected actual)))))
          )))))

   (test-case
    "test-from-row1"

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

       (set-STYLES-styles!
        (*STYLES*)
        (list
         (STYLE
          (BORDER-STYLE "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed")
          #f #f #f #f)))

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml->hash (open-input-string (format "<worksheet><sheetData>~a~a~a</sheetData></worksheet>"
                                                       (file->string row1_file)
                                                       (file->string row2_file)
                                                       (file->string row3_file)
                                                       )))])

            (check-false (hash-has-key? (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) 1))

            (check-true (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1"))

            (from-row xml_hash 1)

            (check-true (hash-has-key? (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) 1))

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "month2")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "month3")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "real")

            (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 1) 5)

            (from-row xml_hash 2)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A2") 201601)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B2") 100)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C2") 110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D2") 1110)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E2") 6.9)

            (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 2) 5)

            (from-row xml_hash 3)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D3") 43360)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E3") 43360)

            (check-equal? (hash-ref (SHEET-STYLE-row->height_map (*CURRENT_SHEET_STYLE*)) 3) 5)

            (from-row xml_hash 4)
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A4"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B4"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C4"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D4"))
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E4"))

            )))
       )))

   (test-case
    "test-from-row2"

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

       (set-STYLES-styles!
        (*STYLES*)
        (list
         (STYLE
          (BORDER-STYLE "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed")
          #f #f #f #f)))

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml->hash (open-input-string (format "<worksheet><sheetData>~a</sheetData></worksheet>"
                                                       (file->string row1_not_from_a1_file)
                                                       )))])

            (check-false (hash-has-key? (SHEET-STYLE-row->style_map (*CURRENT_SHEET_STYLE*)) 1))

            (hash-clear! (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)))

            (from-row xml_hash 1)

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C3") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D3") "month2")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E3") "month3")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F3") "month1")
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "G3") "real")
            )))
       )))
   ))

(run-tests test-row)
