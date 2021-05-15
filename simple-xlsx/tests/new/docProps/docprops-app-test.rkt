#lang racket

(require simple-xml)

(require rackunit/text-ui)
(require rackunit "../../../new/docProps/docprops-app.rkt")

(require "../../../xlsx/xlsx.rkt")
(require "../../../writer.rkt")
(require "../../../lib/lib.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "app.xml")

(define test-docprops-app
  (test-suite
   "test-docprops-app"

   (test-case
    "test-docprops-app"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (add-data-sheet "数据页面" '((1)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      (add-chart-sheet "Chart1" 'LINE "Chart1")
      (add-chart-sheet "Chart4" 'LINE "Chart4")
      
      (printf "~a\n" (docprops-app))

      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (docprops-app))
           (lambda (actual)
             (check-lines? expected actual)))))))))

(run-tests test-docprops-app)
