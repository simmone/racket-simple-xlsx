#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt"
         "../../lib/lib.rkt"
         rackunit/text-ui
         rackunit
         "../../docProps/docprops-app.rkt"
         racket/runtime-path)

(define-runtime-path app_file "app.xml")

(define test-docprops-app
  (test-suite
   "test-docprops-app"

   (test-case
    "test-write-docprops-app"

    (with-xlsx
     (lambda ()
       (add-data-sheet "数据页面" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart4" 'LINE "Chart4" '())

       (dynamic-wind
           (lambda ()
             (write-docprops-app (apply build-path (drop-right (explode-path app_file) 1))))
           (lambda ()
             (call-with-input-file app_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (to-docprops-app))
                  (lambda (actual)
                    (check-lines? expected actual)))))

             (read-docprops-app (apply build-path (drop-right (explode-path app_file) 1)))
             (let ([sheet_list (XLSX-sheet_list (*XLSX*))])
               (check-equal? (length sheet_list) 5)
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
               )
             )
           (lambda ()
             (when (file-exists? app_file) (delete-file app_file)))))))
   ))

(run-tests test-docprops-app)
