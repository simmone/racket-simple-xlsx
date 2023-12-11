#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../../xlsx/xlsx.rkt")
(require "../../../../../../sheet/sheet.rkt")
(require "../../../../../../lib/lib.rkt")

(require"../../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path number_cell_file "number_cell.xml")
(define-runtime-path number_cell_type_n_file "number_cell_type_n.xml")

(define test-cell
  (test-suite
   "test-cell"

   (test-case
    "test-to-cell"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((100.123)))

       (with-sheet
        (lambda ()
          (call-with-input-file number_cell_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-cell 1 1))
               (lambda (actual)
                 (check-lines? expected actual)))))
          )))))

   (test-case
    "test-from-cell"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("" "" "" "" "")))

       (with-sheet
        (lambda ()
          (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "")
          (from-cell
           (xml->hash (open-input-string (format "<worksheet><sheetData><row r=\"1\">~a</row></sheetData></worksheet>"
                                                 (file->string number_cell_file)
                                                 )))
           1 1)
          (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") 100.123)
          )))))

   (test-case
    "test-from-cell-type-n"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("" "" "" "" "")))

       (with-sheet
        (lambda ()
          (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") "")
          (from-cell
           (xml->hash (open-input-string (format "<worksheet><sheetData><row r=\"1\">~a</row></sheetData></worksheet>"
                                                 (file->string number_cell_type_n_file)
                                                 )))
           1 1)
          (check-equal? (hash-ref (DATA-SHEET-cell->value_hash (*CURRENT_SHEET*)) "A1") 100.123)
          )))))

   ))

(run-tests test-cell)
