#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../xl/charts/charts-lib.rkt"
         racket/runtime-path)

(define-runtime-path ser_cat_file "ser_cat.xml")
(define-runtime-path sers_file "sers.xml")

(define test-ser-cat
  (test-suite
   "test-ser-cat"

   (test-case
    "test-ser-cat"

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
       (call-with-input-file ser_cat_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content (to-ser-cat "Sheet1" "A1-D1"))
            (lambda (actual)
              (check-lines? expected actual))))))))))

   (test-case
    "test-from-ser-cat"

    (with-xlsx
     (lambda ()
      (add-chart-sheet "Chart1" 'LINE "Chart1" '())
      (add-chart-sheet "Chart2" 'LINE3D "Chart2" '())
      (add-chart-sheet "Chart3" 'PIE "Chart3" '())
      (add-chart-sheet "Chart4" 'PIE3D "Chart4" '())
      (add-chart-sheet "Chart5" 'BAR "Chart5" '())
      (add-chart-sheet "Chart6" 'BAR3D "Chart6" '())

      (with-sheet-name
       "Chart1"
       (lambda ()
         (check-equal?
          (from-ser-cat
           (xml-port-to-hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:lineChart>~a</c:lineChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file)))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:cat.c:strRef.c:f"
              )
            )
           1)
          '("DataSheet1" . "A1-C1"))))

      (with-sheet-name
       "Chart2"
       (lambda ()
         (check-equal?
          (from-ser-cat
           (xml-port-to-hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:line3DChart>~a</c:line3DChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file)))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:line3DChart.c:ser.c:cat.c:strRef.c:f"
              )
            )
           2)
          '("DataSheet2" . "D1-F1"))))

      (with-sheet-name
       "Chart3"
       (lambda ()
         (check-equal?
          (from-ser-cat
           (xml-port-to-hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:pieChart>~a</c:pieChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file)))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:cat.c:strRef.c:f"
              )
            )
           3)
          '("DataSheet3" . "G1-I1"))))

      (with-sheet-name
       "Chart4"
       (lambda ()
         (check-equal?
          (from-ser-cat
           (xml-port-to-hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:pie3DChart>~a</c:pie3DChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file)))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:pie3DChart.c:ser.c:cat.c:strRef.c:f"
              ))
           4)
          '("DataSheet4" . "J1-L1"))))

      (with-sheet-name
       "Chart5"
       (lambda ()
         (check-equal?
          (from-ser-cat
           (xml-port-to-hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:barChart>~a</c:barChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file)))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:cat.c:strRef.c:f"
              ))
           5)
          '("DataSheet5" . "M1-O1"))))

      (with-sheet-name
       "Chart6"
       (lambda ()
         (check-equal?
          (from-ser-cat
           (xml-port-to-hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:bar3DChart>~a</c:bar3DChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file)))
            '(
              "c:chartSpace.c:chart.c:plotArea.c:bar3DChart.c:ser.c:cat.c:strRef.c:f"
              ))
           6)
          '("DataSheet6" . "P1-R1"))))
      )))

   ))

(run-tests test-ser-cat)
