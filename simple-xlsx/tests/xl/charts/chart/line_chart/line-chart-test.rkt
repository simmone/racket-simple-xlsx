#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../lib/lib.rkt")
(require "../../../../../xl/charts/charts.rkt")
(require "../../../../../xl/charts/lib.rkt")

(require racket/runtime-path)
(define-runtime-path line_chart_file "line_chart.xml")
(define-runtime-path line_3d_chart_file "line_3d_chart.xml")

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
        "LineChart"
        'LINE
        "LineChartExample"
        '(
         ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
         ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
         ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
         ))

       (add-chart-sheet
        "Line3DChart"
        'LINE3D
        "Line3DChartExample"
        '(
         ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
         ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
         ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
         ))

       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       ;; test chart-line
       (with-sheet-name
        "LineChart"
        (lambda ()
          (call-with-input-file line_chart_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-chart-line))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-name
        "Line3DChart"
        (lambda ()
          (call-with-input-file line_3d_chart_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content
                (to-chart-3d-line))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-name
        "LineChart"
        (lambda ()
          (call-with-input-file line_chart_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-chart))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       (with-sheet-name
        "Line3DChart"
        (lambda ()
          (call-with-input-file line_3d_chart_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (to-chart))
               (lambda (actual)
                 (check-lines? expected actual)))))))

       )))

   (test-case
    "test-from-chart"

    (with-xlsx
     (lambda ()
       (add-chart-sheet "LineChart" 'UNKNOWN "" '())
       (add-chart-sheet "Line3DChart" 'UNKNOWN "" '())

       (with-sheet-name
        "LineChart"
        (lambda ()
          (let ([line_chart
                 (from-chart
                  (xml->hash
                   (open-input-string (file->string line_chart_file))))])
          
            (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE)
            (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "LineChartExample")
            (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
              (check-equal? (length sers) 3)
              
              (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))
              (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3"))
              (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))
            )))

       (with-sheet-name
        "Line3DChart"
        (lambda ()
          (let ([line_3d_chart
                 (from-chart
                  (xml->hash
                   (open-input-string (file->string line_3d_chart_file))))])
          
            (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE3D)
            (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "Line3DChartExample")
            (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
              (check-equal? (length sers) 3)
              
              (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))
              (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3"))
              (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))
            ))))))
    ))

  (run-tests test-charts)
