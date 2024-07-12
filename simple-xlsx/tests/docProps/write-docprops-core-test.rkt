#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../lib/lib.rkt"
         rackunit/text-ui
         rackunit
         "../../docProps/docprops-core.rkt"
         racket/runtime-path)

(define-runtime-path core_file "core.xml")

(define test-docprops-core
  (test-suite
   "test-docprops-core"

   (test-case
    "test-write-docprops-core"

    (with-xlsx
     (lambda ()
       (add-data-sheet "数据页面" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart4" 'LINE "Chart4" '())

       (dynamic-wind
           (lambda ()
             (write-docprops-core
              (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")
              (apply build-path (drop-right (explode-path core_file) 1))))
           (lambda ()
             (call-with-input-file core_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (docprops-core (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")))
                  (lambda (actual)
                    (check-lines? expected actual))))))
           (lambda ()
             (when (file-exists? core_file) (delete-file core_file)))))))
   ))

(run-tests test-docprops-core)
