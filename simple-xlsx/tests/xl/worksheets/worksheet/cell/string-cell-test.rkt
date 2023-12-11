#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../style/style.rkt")
(require "../../../../../style/border-style.rkt")
(require "../../../../../style/styles.rkt")
(require "../../../../../style/assemble-styles.rkt")
(require "../../../../../style/set-styles.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../lib/lib.rkt")
(require "../../../../../lib/sheet-lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path string1_cell_file "string1_cell.xml")
(define-runtime-path string2_cell_file "string2_cell.xml")
(define-runtime-path string3_cell_file "string3_cell.xml")
(define-runtime-path string4_cell_file "string4_cell.xml")
(define-runtime-path string5_cell_file "string5_cell.xml")
(define-runtime-path string6_cell_file "string6_cell.xml")

(define test-cell
  (test-suite
   "test-cell"

   (test-case
    "test-to-cell"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real")))

       (squash-shared-strings-map)

       (with-sheet
        (lambda ()
          (set-cell-range-border-style "A1-C1" "all" "0000FF" "thick")

          (strip-styles)

          (call-with-input-file string1_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cell 1 1))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string2_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cell 1 2))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string3_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cell 1 3))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string4_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cell 1 4))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string5_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cell 1 5))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (check-equal? '() (to-cell 1 6)))
        ))))

   (test-case
    "test-from-cell"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("" "" "" "" "")))

       (set-STYLES-styles!
        (*STYLES*)
        (list
         (STYLE
          (BORDER-STYLE "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed")
          #f #f #f #f)))

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month1" 0)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 0 "month1")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month2" 1)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 1 "month2")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month3" 2)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 2 "month3")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "month4" 3)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 3 "month4")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "real" 4)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 4 "real")

       (with-sheet
        (lambda ()

          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1"))

          (let ([xml_hash
                 (xml->hash
                  (open-input-string
                   (format "<worksheet><sheetData><row r=\"1\">~a~a~a~a~a~a</row></sheetData></worksheet>"
                           (file->string string5_cell_file)
                           (file->string string2_cell_file)
                           (file->string string3_cell_file)
                           (file->string string4_cell_file)
                           (file->string string1_cell_file)
                           (file->string string6_cell_file)
                           )))])

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "")
            (from-cell xml_hash 1 5)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "month1")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "")
            (from-cell xml_hash 1 2)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "month2")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "")
            (from-cell xml_hash 1 3)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "month3")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "")
            (from-cell xml_hash 1 4)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "month4")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "")
            (from-cell xml_hash 1 1)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "real")

            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F1"))
            (from-cell xml_hash 1 6)
            (check-false (hash-has-key? (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F1"))
            )

          (check-true (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "A1"))
          (check-true (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "B1"))
          (check-true (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "C1"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "D1"))
          (check-false (hash-has-key? (SHEET-STYLE-cell->style_map (*CURRENT_SHEET_STYLE*)) "E1"))
          )))))

   ))

(run-tests test-cell)
