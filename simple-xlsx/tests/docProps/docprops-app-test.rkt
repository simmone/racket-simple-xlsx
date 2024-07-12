#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt"
         "../../lib/lib.rkt"
         rackunit/text-ui
         rackunit
         "../../docProps/docprops-app.rkt"
         racket/runtime-path)

(define-runtime-path app1_file "app1_test.xml")
(define-runtime-path app2_file "app2_test.xml")
(define-runtime-path app3_file "app3_test.xml")

(define test-docprops-app
  (test-suite
   "test-docprops-app"

   (test-case
    "test-write-docprops-app1"

    (with-xlsx
     (lambda ()
       (add-data-sheet "数据页面" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart4" 'LINE "Chart4" '())
       (add-chart-sheet "Chart5" 'LINE "Chart5" '())

       (call-with-input-file app1_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml (to-docprops-app))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-write-docprops-app2"

    (with-xlsx
     (lambda ()
       (add-data-sheet "数据页面" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (call-with-input-file app2_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml (to-docprops-app))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-read-docprops-app"

    (with-xlsx
     (lambda ()
       (add-data-sheet "1" '((1)))
       (add-data-sheet "2" '((1)))
       (add-data-sheet "3" '((1)))
       (add-chart-sheet "4" 'LINE "Chart1" '())
       (add-chart-sheet "5" 'LINE "Chart4" '())
       (add-chart-sheet "6" 'LINE "Chart5" '())

       (from-docprops-app app1_file)

       (let ([sheet_list (XLSX-sheet_list (*XLSX*))])
         (check-equal? (length sheet_list) 6)
         (check-true (DATA-SHEET? (list-ref sheet_list 0)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 0)) "数据页面")
         (check-true (DATA-SHEET? (list-ref sheet_list 1)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 1)) "Sheet2")
         (check-true (DATA-SHEET? (list-ref sheet_list 2)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 2)) "Sheet3")
         (check-true (CHART-SHEET? (list-ref sheet_list 3)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 3)) "Chart1")
         (check-true (CHART-SHEET? (list-ref sheet_list 4)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 4)) "Chart4")
         (check-true (CHART-SHEET? (list-ref sheet_list 5)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 5)) "Chart5")
         )))
    )

   (test-case
    "test-read-docprops-app-not-exist"

    (with-xlsx
     (lambda ()
       (add-data-sheet "1" '((1)))
       (add-data-sheet "2" '((1)))
       (add-data-sheet "3" '((1)))
       (add-chart-sheet "4" 'LINE "Chart1" '())
       (add-chart-sheet "5" 'LINE "Chart4" '())
       (add-chart-sheet "6" 'LINE "Chart5" '())

       (from-docprops-app app3_file)

       (let ([sheet_list (XLSX-sheet_list (*XLSX*))])
         (check-equal? (length sheet_list) 6)
         (check-true (DATA-SHEET? (list-ref sheet_list 0)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 0)) "1")
         (check-true (DATA-SHEET? (list-ref sheet_list 1)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 1)) "2")
         (check-true (DATA-SHEET? (list-ref sheet_list 2)))
         (check-equal? (DATA-SHEET-sheet_name (list-ref sheet_list 2)) "3")
         (check-true (CHART-SHEET? (list-ref sheet_list 3)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 3)) "4")
         (check-true (CHART-SHEET? (list-ref sheet_list 4)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 4)) "5")
         (check-true (CHART-SHEET? (list-ref sheet_list 5)))
         (check-equal? (CHART-SHEET-sheet_name (list-ref sheet_list 5)) "6")
         )))
    )

   ))

(run-tests test-docprops-app)
