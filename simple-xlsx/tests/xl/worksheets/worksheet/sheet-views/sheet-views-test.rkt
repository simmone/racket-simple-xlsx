#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../style/style.rkt"
         "../../../../../style/set-styles.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../xl/worksheets/worksheet.rkt"
         racket/runtime-path)

(define-runtime-path sheet_views1_file "sheet_views1.xml")
(define-runtime-path sheet_views2_file "sheet_views2.xml")
(define-runtime-path sheet_views3_file "sheet_views3.xml")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-sheet-views"

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
          (set-freeze-row-col-range 1 0)))

       (with-sheet-ref
        2
        (lambda ()
          (set-freeze-row-col-range 1 1)))

       (with-sheet-ref
        0
        (lambda ()
          (call-with-input-file sheet_views1_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-sheet-views))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        1
        (lambda ()
          (call-with-input-file sheet_views2_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-sheet-views))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-ref
        2
        (lambda ()
          (call-with-input-file sheet_views3_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-sheet-views))
               (lambda (actual)
                 (check-lines? expected actual)))))))
       )))
   ))

(run-tests test-worksheet)
