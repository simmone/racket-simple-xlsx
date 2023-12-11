#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../../xlsx/xlsx.rkt")
(require "../../../../../sheet/sheet.rkt")
(require "../../../../../lib/lib.rkt")
(require"../../../../../xl/charts/lib.rkt")

(require racket/runtime-path)
(define-runtime-path ser_head_file "ser_head.xml")
(define-runtime-path sers_file "sers.xml")

(define test-ser-head
  (test-suite
   "test-ser-head"

   (test-case
    "test-ser-head"

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
       (call-with-input-file ser_head_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content (to-ser-head 0 "Cat"))
            (lambda (actual)
              (check-lines? expected actual)))))))
      )))

   (test-case
    "test-from-ser-head"

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
          (from-ser-head
           (xml->hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:lineChart>~a</c:lineChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file))))
           1)
          "Ser1")))

      (with-sheet-name
       "Chart2"
       (lambda ()
         (check-equal?
          (from-ser-head
           (xml->hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:line3DChart>~a</c:line3DChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file))))
           2)
          "Ser2")))

      (with-sheet-name
       "Chart3"
       (lambda ()
         (check-equal?
          (from-ser-head
           (xml->hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:pieChart>~a</c:pieChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file))))
           3)
          "Ser3")))

      (with-sheet-name
       "Chart4"
       (lambda ()
         (check-equal?
          (from-ser-head
           (xml->hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:pie3DChart>~a</c:pie3DChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file))))
           4)
          "Ser4")))

      (with-sheet-name
       "Chart5"
       (lambda ()
         (check-equal?
          (from-ser-head
           (xml->hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:barChart>~a</c:barChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file))))
           5)
          "Ser5")))

      (with-sheet-name
       "Chart6"
       (lambda ()
         (check-equal?
          (from-ser-head
           (xml->hash
            (open-input-string
             (format "<c:chartSpace><c:chart><c:plotArea><c:bar3DChart>~a</c:bar3DChart></c:plotArea></c:chart></c:chartSpace>"
                     (file->string sers_file))))
           6)
          "Ser6")))
      )))
   ))

(run-tests test-ser-head)
