#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../xl/charts/charts-lib.rkt"
         racket/runtime-path)

(define-runtime-path ser_file "ser.xml")
(define-runtime-path sers_file "sers.xml")

(define test-ser
  (test-suite
   "test-ser"

   (test-case
    "test-ser"

    (with-xlsx
     (lambda ()
       (add-data-sheet "DataSheet1"
                       '(("month1" "201601" "201602" "201603" "real") (201601 100 300 200 6.9)))
       (add-data-sheet "DataSheet2"
                       '(("month1" "201601" "201602" "201603" "real") (201601 100 300 200 6.9)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (call-with-input-file ser_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content (to-ser 0 '("CAT" "DataSheet1" "B1-D1" "DataSheet2" "B2-D2")))
               (lambda (actual)
                 (check-lines? expected actual))))))))))

   (test-case
    "test-from-sers"

    (with-xlsx
     (lambda ()
       (add-chart-sheet "Chart1" 'LINE "" '())
       (add-chart-sheet "Chart3" 'PIE "" '())
       (add-chart-sheet "Chart5" 'BAR "" '())

       (with-sheet-name
        "Chart1"
        (lambda ()
          (check-equal?
           (from-sers
            (xml-port-to-hash
             (open-input-string
              (format "<c:chartSpace><c:chart><c:plotArea><c:lineChart>~a</c:lineChart></c:plotArea></c:chart></c:chartSpace>"
                      (file->string sers_file)))
             '(
               "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:tx.c:v"
               "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:cat.c:strRef.c:f"
               "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:val.c:numRef.c:f"
               )))
           '(("CAT" "DataSheet1" "A1-C1" "DataSheet12" "D1-F1")
             ("DOG" "DataSheet2" "B1-D1" "DataSheet22" "E1-G1")
             ("MOUSE" "DataSheet3" "C1-E1" "DataSheet32" "F1-H1")))))

       (with-sheet-name
        "Chart1"
        (lambda ()
          (check-equal?
           (from-ser
            (xml-port-to-hash
             (open-input-string
              (format "<c:chartSpace><c:chart><c:plotArea><c:lineChart>~a</c:lineChart></c:plotArea></c:chart></c:chartSpace>"
                      (file->string sers_file)))
             '(
               "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:tx.c:v"
               "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:cat.c:strRef.c:f"
               "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:val.c:numRef.c:f"
               ))
            1)
           '("CAT" "DataSheet1" "A1-C1" "DataSheet12" "D1-F1"))))

       (with-sheet-name
        "Chart3"
        (lambda ()
          (check-equal?
            (from-ser
             (xml-port-to-hash
              (open-input-string
               (format "<c:chartSpace><c:chart><c:plotArea><c:pieChart>~a</c:pieChart></c:plotArea></c:chart></c:chartSpace>"
                       (file->string sers_file)))
              '(
               "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:tx.c:v"
               "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:cat.c:strRef.c:f"
               "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:val.c:numRef.c:f"
               )
              )
             2)
            '("DOG" "DataSheet2" "B1-D1" "DataSheet22" "E1-G1"))))

       (with-sheet-name
        "Chart5"
        (lambda ()
          (check-equal?
            (from-ser
             (xml-port-to-hash
              (open-input-string
               (format "<c:chartSpace><c:chart><c:plotArea><c:barChart>~a</c:barChart></c:plotArea></c:chart></c:chartSpace>"
                       (file->string sers_file)))
              '(
                "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:tx.c:v"
                "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:cat.c:strRef.c:f"
                "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:val.c:numRef.c:f"
               ))
             3)
            '("MOUSE" "DataSheet3" "C1-E1" "DataSheet32" "F1-H1"))))
       )))

   ))

(run-tests test-ser)

