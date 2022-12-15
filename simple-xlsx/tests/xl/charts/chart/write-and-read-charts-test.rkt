#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")
(require"../../../../xl/charts/charts.rkt")
(require"../../../../xl/charts/lib.rkt")

(require racket/runtime-path)
(define-runtime-path chart1_file "chart1.xml")
(define-runtime-path chart2_file "chart2.xml")
(define-runtime-path chart3_file "chart3.xml")
(define-runtime-path chart4_file "chart4.xml")
(define-runtime-path chart5_file "chart5.xml")
(define-runtime-path chart6_file "chart6.xml")

(define test-write-charts
  (test-suite
   "test-write-charts"

   (test-case
    "test-write-charts"

    (with-xlsx
     (lambda ()
      (add-data-sheet "DataSheet"
                      '(
                        ("month1" "201601" "201602" "201603" "real")
                        (201601 100 300 200 6.9)
                        (201601 200 400 300 6.9)
                        (201601 300 500 400 6.9)
                        ))

      (add-chart-sheet
       "LineChart"
       'LINE
       "LineChartExample"
       '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
        ))
      
      (add-data-sheet "Sheet2" '(("none")))

      (add-chart-sheet
       "Line3DChart"
       'LINE3D
       "Line3DChartExample"
       '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
        ))

      (add-chart-sheet
       "PieChart"
       'PIE
       "PieChartExample"
       '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ))

      (add-chart-sheet
       "Pie3DChart"
       'PIE3D
       "Pie3DChartExample"
       '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ))

      (add-chart-sheet
       "BarChart"
       'BAR
       "BarChartExample"
       '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
        ))

      (add-chart-sheet
       "Bar3DChart"
       'BAR3D
       "Bar3DChartExample"
       '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")
        ))

      (dynamic-wind
          (lambda ()
            (write-charts (apply build-path (drop-right (explode-path chart1_file) 1))))
          (lambda ()
             ;; line chart
            (with-sheet-ref
             1
             (lambda ()
               (call-with-input-file chart1_file
                 (lambda (expected)
                   (call-with-input-string
                    (lists->xml (to-chart))
                    (lambda (actual)
                      (check-lines? expected actual)))))))

             ;; line 3d chart
            (with-sheet-ref
             3
             (lambda ()
               (call-with-input-file chart2_file
                 (lambda (expected)
                   (call-with-input-string
                    (lists->xml (to-chart))
                    (lambda (actual)
                      (check-lines? expected actual)))))))

             ;; pie chart
            (with-sheet-ref
             4
             (lambda ()
               (call-with-input-file chart3_file
                 (lambda (expected)
                   (call-with-input-string
                    (lists->xml (to-chart)) 
                    (lambda (actual)
                      (check-lines? expected actual)))))))

             ;; pie 3d chart
            (with-sheet-ref
             5
             (lambda ()
               (call-with-input-file chart4_file
                 (lambda (expected)
                   (call-with-input-string
                    (lists->xml (to-chart))
                    (lambda (actual)
                      (check-lines? expected actual)))))))

             ;; bar chart
            (with-sheet-ref
             6
             (lambda ()
               (call-with-input-file chart5_file
                 (lambda (expected)
                   (call-with-input-string
                    (lists->xml (to-chart))
                    (lambda (actual)
                      (check-lines? expected actual)))))))

             ;; bar 3d chart
            (with-sheet-ref
             7
             (lambda ()
               (call-with-input-file chart6_file
                 (lambda (expected)
                   (call-with-input-string
                    (lists->xml (to-chart))
                    (lambda (actual)
                      (check-lines? expected actual)))))))
            
            (set-XLSX-sheet_list! (*XLSX*) '())
            
            (check-equal? (length (XLSX-sheet_list (*XLSX*))) 0)

            (add-data-sheet "DataSheet" '(("none")))
            (add-chart-sheet "LineChart" 'UNKNOWN "" '())
            (add-chart-sheet "Line3DChart" 'UNKNOWN "" '())
            (add-chart-sheet "PieChart" 'UNKNOWN "" '())
            (add-chart-sheet "Pie3DChart" 'UNKNOWN "" '())
            (add-chart-sheet "BarChart" 'UNKNOWN "" '())
            (add-chart-sheet "Bar3DChart" 'UNKNOWN "" '())

            (read-charts (apply build-path (drop-right (explode-path chart1_file) 1)))

            (with-sheet-name
             "LineChart"
             (lambda ()
               (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE)
               (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "LineChartExample")
               (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
                 (check-equal? (length sers) 3)
              
                 (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))
                 (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3"))
                 (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))
               ))

            (with-sheet-name
             "Line3DChart"
             (lambda ()
               (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE3D)
               (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "Line3DChartExample")
               (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
                 (check-equal? (length sers) 3)
                 
                 (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))
                 (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3"))
                 (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))
               ))

            (with-sheet-name
             "BarChart"
             (lambda ()
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
               (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'BAR3D)
               (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "Bar3DChartExample")
               (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
                 (check-equal? (length sers) 3)
                 
                 (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2"))
                 (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3"))
                 (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))
               ))

            (with-sheet-name
             "PieChart"
             (lambda ()
               (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE)
               (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "PieChartExample")
               (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
                 (check-equal? (length sers) 1)
                 
                 (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")))
               ))

            (with-sheet-name
             "Pie3DChart"
             (lambda ()
               (check-eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE3D)
               (check-equal? (CHART-SHEET-topic (*CURRENT_SHEET*)) "Pie3DChartExample")
               (let ([sers (CHART-SHEET-serial (*CURRENT_SHEET*))])
                 (check-equal? (length sers) 1)
                 
                 (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")))
               ))
            )
          (lambda ()
            (when (file-exists? chart1_file) (delete-file chart1_file))
            (when (file-exists? chart2_file) (delete-file chart2_file))
            (when (file-exists? chart3_file) (delete-file chart3_file))
            (when (file-exists? chart4_file) (delete-file chart4_file))
            (when (file-exists? chart5_file) (delete-file chart5_file))
            (when (file-exists? chart6_file) (delete-file chart6_file))))
        )))
    ))

(run-tests test-write-charts)
