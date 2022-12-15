#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../lib/lib.rkt")

(require"../../../../../xl/worksheets/worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path work_sheet_head_file "work_sheet_head.xml")
(define-runtime-path work_sheet_not_start_from_A1_head_file "work_sheet_not_start_from_A1_head.xml")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-work-sheet-head"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (call-with-input-file work_sheet_head_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-work-sheet-head))
               (lambda (actual)
                 (check-lines? expected actual)))))))
       ))

    (with-xlsx
     (lambda ()
       (add-data-sheet "S1" '((1)))

       (with-sheet
        (lambda ()

          (from-work-sheet-head (xml->hash work_sheet_head_file))

          (check-equal? (DATA-SHEET-dimension (*CURRENT_SHEET*)) "A1:E2")))))

    (with-xlsx
     (lambda ()
       (add-data-sheet "S1" '((1)))

       (with-sheet
        (lambda ()

          (from-work-sheet-head (xml->hash work_sheet_not_start_from_A1_head_file))

          (check-equal? (DATA-SHEET-dimension (*CURRENT_SHEET*)) "B1:E2")))))
    )
   ))

(run-tests test-worksheet)
