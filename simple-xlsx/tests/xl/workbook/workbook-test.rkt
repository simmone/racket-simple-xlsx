#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../sheet/sheet.rkt"
         "../../../lib/lib.rkt"
         "../../../xl/workbook.rkt"
         racket/runtime-path)

(define-runtime-path workbook_file "workbook_test.xml")

(define test-workbook
  (test-suite
   "test-workbook"

   (test-case
    "test-to-workbook"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("1")))
       (add-data-sheet "Sheet3" '((1)))
       (add-data-sheet "Sheet5" '((1)))
       (add-chart-sheet "Chart2" 'LINE "Chart1" '())
       (add-chart-sheet "Chart4" 'LINE "Chart2" '())
       (add-chart-sheet "Chart6" 'LINE "Chart3" '())

       (call-with-input-file workbook_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml (to-workbook))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-workbook"

    (with-xlsx
     (lambda ()
       (add-data-sheet "S1" '(("1")))
       (add-data-sheet "S2" '((1)))
       (add-data-sheet "S3" '((1)))
       (add-chart-sheet "C1" 'LINE "Chart1" '())
       (add-chart-sheet "C2" 'LINE "Chart2" '())
       (add-chart-sheet "C3" 'LINE "Chart3" '())

       (from-workbook workbook_file)

       (let ([sheet_list (XLSX-sheet_list (*XLSX*))])
         (check-equal? (length sheet_list) 6)
         (check-true (DATA-SHEET? (list-ref sheet_list 0)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 0)) "Sheet1")
         (check-true (DATA-SHEET? (list-ref sheet_list 1)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 1)) "Sheet3")
         (check-true (DATA-SHEET? (list-ref sheet_list 2)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 2)) "Sheet5")
         (check-true (CHART-SHEET? (list-ref sheet_list 3)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 3)) "Chart2")
         (check-true (CHART-SHEET? (list-ref sheet_list 4)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 4)) "Chart4")
         (check-true (CHART-SHEET? (list-ref sheet_list 5)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 5)) "Chart6")
         )))
    )
   ))

(run-tests test-workbook)
