#lang racket

(require fast-xml
         rackunit/text-ui rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/worksheets/_rels/worksheets-rels.rkt"
         racket/runtime-path)

(define-runtime-path sheet1_rels_file "sheet1.xml.rels")
(define-runtime-path sheet2_rels_file "sheet2.xml.rels")
(define-runtime-path sheet3_rels_file "sheet3.xml.rels")

(define test-worksheets-rels
  (test-suite
   "test-worksheets-rels"

   (test-case
    "test-write-worksheets-rels"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("1")))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (dynamic-wind
           (lambda ()
             (write-worksheets-rels (apply build-path (drop-right (explode-path sheet1_rels_file) 1))))
           (lambda ()
             (call-with-input-file sheet1_rels_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (worksheets-rels 1))
                  (lambda (actual)
                    (check-lines? expected actual)))))

             (call-with-input-file sheet2_rels_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (worksheets-rels 2))
                  (lambda (actual)
                    (check-lines? expected actual)))))

             (call-with-input-file sheet3_rels_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (worksheets-rels 3))
                  (lambda (actual)
                    (check-lines? expected actual))))))
           (lambda ()
             (when (file-exists? sheet1_rels_file) (delete-file sheet1_rels_file))
             (when (file-exists? sheet2_rels_file) (delete-file sheet2_rels_file))
             (when (file-exists? sheet3_rels_file) (delete-file sheet3_rels_file))
             )))))
   ))

(run-tests test-worksheets-rels)
