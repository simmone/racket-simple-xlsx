#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../xl/charts/charts.rkt"
         "../../../../../xl/charts/charts-lib.rkt"
         racket/runtime-path)

(define-runtime-path bar_chart_file "bar_chart.xml")
(define-runtime-path bar_3d_chart_file "bar_3d_chart.xml")

(define test-charts
  (test-suite
   "test-charts"

   (test-case
    "test-charts"

    (with-xlsx
     (lambda ()
      (add-data-sheet "DataSheet"
                      '(
                        ("month1" "201601" "201602" "201603" "real")
                        (201601 100 300 200 6.9)
                        (201601 200 400 300 6.9)
                        (201601 300 500 400 6.9)
                        ))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))

      (add-chart-sheet
       "BarChart"
       'BAR
       "BarChartExample"
       (list
        '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
        ))

      (add-chart-sheet
       "Bar3DChart"
       'BAR3D
       "Bar3DChartExample"
       (list
        '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
        ))

      (add-chart-sheet "Chart3" 'BAR "Chart3" '())

      (with-sheet-name
       "BarChart"
       (lambda ()
         (call-with-input-file bar_chart_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart-bar))
              (lambda (actual)
                (check-lines? expected actual)))))))

      (with-sheet-name
       "Bar3DChart"
       (lambda ()
         (call-with-input-file bar_3d_chart_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart-3d-bar))
              (lambda (actual)
                (check-lines? expected actual)))))))

      (with-sheet-name
       "BarChart"
       (lambda ()
         (call-with-input-file bar_chart_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart))
              (lambda (actual)
                (check-lines? expected actual)))))))

      (with-sheet-name
       "Bar3DChart"
       (lambda ()
         (call-with-input-file bar_3d_chart_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart))
              (lambda (actual)
                (check-lines? expected actual)))))))
      )))

   (test-case
    "test-from-chart"

    (with-xlsx
     (lambda ()
       (add-chart-sheet "BarChart" 'UNKNOWN "" '())
       (add-chart-sheet "Bar3DChart" 'UNKNOWN "" '())

       (with-sheet-name
        "BarChart"
        (lambda ()
          (from-chart
           (xml-port-to-hash
            (open-input-string (file->string bar_chart_file))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:tx.c:v"
              "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:cat.c:strRef.c:f"
              "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:val.c:numRef.c:f"
              "c:chartSpace.c:chart.c:plotArea.c:barChart"
              "c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.a:r.a:t"
              )
            ))

          (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'BAR)

          (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "BarChartExample")

          (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
            (check-equal? (length sers) 3)
            
            (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))
            (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3"))
            (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))
          ))

       (with-sheet-name
        "Bar3DChart"
        (lambda ()
          (let ([bar_3d_chart
                 (from-chart
                  (xml-file-to-hash
                   bar_3d_chart_file
                   '(
                     "c:chartSpace.c:chart.c:plotArea.c:bar3DChart.c:ser.c:tx.c:v"
                     "c:chartSpace.c:chart.c:plotArea.c:bar3DChart.c:ser.c:cat.c:strRef.c:f"
                     "c:chartSpace.c:chart.c:plotArea.c:bar3DChart.c:ser.c:val.c:numRef.c:f"
                     "c:chartSpace.c:chart.c:plotArea.c:bar3DChart"
                     "c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.a:r.a:t"
                     )))])

            (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'BAR3D)
            (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "Bar3DChartExample")
            (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
              (check-equal? (length sers) 3)

              (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))
              (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3"))
              (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))
            ))))))

   ))

(run-tests test-charts)
