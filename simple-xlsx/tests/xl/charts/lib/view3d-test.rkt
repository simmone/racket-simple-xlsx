#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/charts/lib.rkt")

(require racket/runtime-path)
(define-runtime-path view3d_file "view3d.xml")

(define test-view3d
  (test-suite
   "test-view3d"

   (test-case
    "test-view3d"

    (with-xlsx
     (lambda ()
       (add-data-sheet "DataSheet"
                       '(("month1" "201601" "201602" "201603" "real") (201601 100 300 200 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (with-sheet
        (lambda ()
          (call-with-input-file view3d_file
            (lambda (expected)
              (call-with-input-string
               (lists->xml_content (view3d))
               (lambda (actual)
                 (check-lines? expected actual))))))))))))

(run-tests test-view3d)

