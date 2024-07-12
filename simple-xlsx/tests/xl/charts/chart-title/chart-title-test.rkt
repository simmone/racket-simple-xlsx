#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/charts/charts.rkt"
         racket/runtime-path)

(define-runtime-path chart_title_file "chart_title.xml")

(define test-charts
  (test-suite
   "test-charts"

   (test-case
    "test-chart-title"

    (with-xlsx
     (lambda ()
      (add-data-sheet "Sheet1"
                      '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      (add-chart-sheet "Chart1" 'LINE "Chart1" '())
      (add-chart-sheet "Chart2" 'LINE "Chart2" '())
      (add-chart-sheet "Chart3" 'LINE "Chart3" '())

      (with-sheet-ref
       0
       (lambda ()
         (call-with-input-file chart_title_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart-title "Chart1"))
              (lambda (actual)
                (check-lines? expected actual)))))))))

    (with-xlsx
     (lambda ()
      (add-data-sheet "Sheet1"
                      '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      (add-chart-sheet "Chart1" 'LINE "" '())
      (add-chart-sheet "Chart2" 'LINE "" '())
      (add-chart-sheet "Chart3" 'LINE "" '())

      (with-sheet-ref
       3
       (lambda ()
         (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "")

         (from-chart-title
           (xml-port-to-hash
            (open-input-string (format "<c:chartSpace><c:chart>~a</c:chart></c:chartSpace>" (file->string chart_title_file)))
            '(
              "c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.a:r.a:t"
              )
            ))

         (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "Chart1")))))
    )
   ))

(run-tests test-charts)
