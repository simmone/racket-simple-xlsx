#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../../xlsx/xlsx.rkt"
         "../../../../../sheet/sheet.rkt"
         "../../../../../lib/lib.rkt"
         "../../../../../xl/charts/charts-lib.rkt"
         racket/runtime-path)

(define-runtime-path sers_file "sers.xml")

(define test-sers
  (test-suite
   "test-sers"

   (test-case
    "test-to-sers"

    (with-xlsx
     (lambda ()
       (add-data-sheet "DataSheet"
                       '(
                         ("month1" "201601" "201602" "201603" "real")
                         ))
       (add-data-sheet "Sheet2"
                       '(
                         (201601 100 300 200 6.9)
                         (201601 200 400 300 6.9)
                         (201601 300 500 400 6.9)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (call-with-input-file sers_file
            (lambda (expected)
              (call-with-input-string
               (lists-to-xml_content
                (append
                 '("c:lineChart")
                 (to-sers
                  '(
                   ("CAT" "DataSheet" "B1-D1" "Sheet2" "B1-D1")
                   ("Puma" "DataSheet" "B1-D1" "Sheet2" "B2-D2")
                   ("Brooks" "DataSheet" "B1-D1" "Sheet2" "B3-D3")
                   ))))
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
          (let ([sers
                 (from-sers
                  (xml-port-to-hash
                   (open-input-string
                    (format "<c:chartSpace><c:chart><c:plotArea>~a</c:plotArea></c:chart></c:chartSpace>"
                            (file->string sers_file)))
                   '(
                     "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:tx.c:v"
                     "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:cat.c:strRef.c:f"
                     "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:val.c:numRef.c:f"
                     )))])
                
            (check-equal? (length sers) 3)

            (check-equal? (list-ref sers 0) '("CAT" "DataSheet" "B1-D1" "Sheet2" "B1-D1"))
            (check-equal? (list-ref sers 1) '("Puma" "DataSheet" "B1-D1" "Sheet2" "B2-D2"))
            (check-equal? (list-ref sers 2) '("Brooks" "DataSheet" "B1-D1" "Sheet2" "B3-D3"))))))))
   ))

(run-tests test-sers)
