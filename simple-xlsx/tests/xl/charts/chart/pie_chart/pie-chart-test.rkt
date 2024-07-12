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

(define-runtime-path pie_chart_file "pie_chart.xml")
(define-runtime-path pie_3d_chart_file "pie_3d_chart.xml")

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
       "PieChart"
       'PIE
       "PieChartExample"
       (list
        '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ))

      (add-chart-sheet
       "Pie3DChart"
       'PIE3D
       "Pie3DChartExample"
       (list
        '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ))

      (add-chart-sheet "Chart3" 'PIE "Chart3" '())

      (with-sheet-name
       "PieChart"
       (lambda ()
         (call-with-input-file pie_chart_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart-pie))
              (lambda (actual)
                (check-lines? expected actual)))))))

      (with-sheet-name
       "Pie3DChart"
       (lambda ()
         (call-with-input-file pie_3d_chart_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart-3d-pie))
              (lambda (actual)
                (check-lines? expected actual)))))))

      (with-sheet-name
       "PieChart"
       (lambda ()
         (call-with-input-file pie_chart_file
           (lambda (expected)
             (call-with-input-string
              (lists-to-xml_content (to-chart))
              (lambda (actual)
                (check-lines? expected actual)))))))

      (with-sheet-name
       "Pie3DChart"
       (lambda ()
         (call-with-input-file pie_3d_chart_file
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
       (add-chart-sheet "PieChart" 'UNKNOWN "" '())
       (add-chart-sheet "Pie3DChart" 'UNKNOWN "" '())

       (with-sheet-name
        "PieChart"
        (lambda ()
          (from-chart
           (xml-port-to-hash
            (open-input-string (file->string pie_chart_file))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:tx.c:v"
              "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:cat.c:strRef.c:f"
              "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:val.c:numRef.c:f"
              "c:chartSpace.c:chart.c:plotArea.c:pieChart"
              "c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.a:r.a:t"
              )))

            (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE)
            (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "PieChartExample")
            (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
              (check-equal? (length sers) 1)

              (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")))
            ))

       (with-sheet-name
        "Pie3DChart"
        (lambda ()
          (from-chart
           (xml-port-to-hash
            (open-input-string (file->string pie_3d_chart_file))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:pie3DChart.c:ser.c:tx.c:v"
              "c:chartSpace.c:chart.c:plotArea.c:pie3DChart.c:ser.c:cat.c:strRef.c:f"
              "c:chartSpace.c:chart.c:plotArea.c:pie3DChart.c:ser.c:val.c:numRef.c:f"
              "c:chartSpace.c:chart.c:plotArea.c:pie3DChart"
              "c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.a:r.a:t"
              )))

            (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE3D)
            (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "Pie3DChartExample")
            (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
              (check-equal? (length sers) 1)

              (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")))
            )))))

    ))

(run-tests test-charts)
