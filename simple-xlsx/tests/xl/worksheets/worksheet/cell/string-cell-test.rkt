#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../style/style.rkt"
         "../../../../../style/border-style.rkt"
         "../../../../../style/styles.rkt"
         "../../../../../style/assemble-styles.rkt"
         "../../../../../style/set-styles.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../lib/sheet-lib.rkt"
         "../../../../../xl/worksheets/worksheet.rkt"
         racket/runtime-path)

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
               (lists-to-xml_content (to-cell 1 1))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string2_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-cell 1 2))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string3_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-cell 1 3))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string4_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-cell 1 4))
               (lambda (actual)
                 (check-lines? expected actual)))))

          (call-with-input-file string5_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-cell 1 5))
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
                 (xml-port-to-hash
                  (open-input-string
                   (format "<worksheet><sheetData><row r=\"1\">~a~a~a~a~a~a</row></sheetData></worksheet>"
                           (file->string string5_cell_file)
                           (file->string string2_cell_file)
                           (file->string string3_cell_file)
                           (file->string string4_cell_file)
                           (file->string string1_cell_file)
                           (file->string string6_cell_file)
                           ))
                  '(
                    "worksheet.sheetData.row.c.r"
                    "worksheet.sheetData.row.c.s"
                    "worksheet.sheetData.row.c.t"
                    "worksheet.sheetData.row.c.v"
                    )
                  )])

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

   (test-case
    "test-from-special-char-cell"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("" "" "" "" "" "" "" "" "")))

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "<test>" 0)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 0 "<test>")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "<foo> " 1)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 1 "<foo> ")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) " <baz>" 2)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 2 " <baz>")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "< bar>" 3)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 3 "< bar>")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "< fro >" 4)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 4 "< fro >")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "<bas >" 5)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 5 "<bas >")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "<maybe" 6)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 6 "<maybe")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "<< not >>" 7)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 7 "<< not >>")

       (hash-set! (XLSX-shared_string->index_map (*XLSX*)) "show>" 8)
       (hash-set! (XLSX-shared_index->string_map (*XLSX*)) 8 "show>")

       (with-sheet
        (lambda ()
          (let ([xml_hash
                 (xml-port-to-hash
                  (open-input-string
                   (format "<worksheet><sheetData><row r=\"1\">~a~a~a~a~a~a~a~a~a</row></sheetData></worksheet>"
                           "<c r=\"A1\" t=\"s\"><v>0</v></c>"
                           "<c r=\"B1\" t=\"s\"><v>1</v></c>"
                           "<c r=\"C1\" t=\"s\"><v>2</v></c>"
                           "<c r=\"D1\" t=\"s\"><v>3</v></c>"
                           "<c r=\"E1\" t=\"s\"><v>4</v></c>"
                           "<c r=\"F1\" t=\"s\"><v>5</v></c>"
                           "<c r=\"G1\" t=\"s\"><v>6</v></c>"
                           "<c r=\"H1\" t=\"s\"><v>7</v></c>"
                           "<c r=\"I1\" t=\"s\"><v>8</v></c>"
                           ))
                  '(
                    "worksheet.sheetData.row.c.r"
                    "worksheet.sheetData.row.c.s"
                    "worksheet.sheetData.row.c.t"
                    "worksheet.sheetData.row.c.v"
                    ))])

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "")
            (from-cell xml_hash 1 1)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "<test>")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "")
            (from-cell xml_hash 1 2)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "B1") "<foo> ")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") "")
            (from-cell xml_hash 1 3)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "C1") " <baz>")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "")
            (from-cell xml_hash 1 4)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "D1") "< bar>")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "")
            (from-cell xml_hash 1 5)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "E1") "< fro >")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F1") "")
            (from-cell xml_hash 1 6)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "F1") "<bas >")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "G1") "")
            (from-cell xml_hash 1 7)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "G1") "<maybe")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "H1") "")
            (from-cell xml_hash 1 8)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "H1") "<< not >>")

            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "I1") "")
            (from-cell xml_hash 1 9)
            (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "I1") "show>")
            )
          )))))

   ))

(run-tests test-cell)
