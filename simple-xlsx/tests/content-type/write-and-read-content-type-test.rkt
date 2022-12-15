#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")
(require "../../lib/lib.rkt")

(require rackunit/text-ui rackunit)

(require"../../content-type.rkt")

(require racket/runtime-path)
(define-runtime-path content_type_file "[Content_Types].xml")

(define test-content-type
  (test-suite
   "test-content-type"

   (test-case
    "test-content-type"

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
             (write-content-type (apply build-path (drop-right (explode-path content_type_file) 1))))
           (lambda ()
             (call-with-input-file content_type_file
               (lambda (expected)
                 (call-with-input-string
                  (lists->xml (to-content-type))
                  (lambda (actual)
                    (check-lines? expected actual)))))

             (with-xlsx
              (lambda ()
                (read-content-type (apply build-path (drop-right (explode-path content_type_file) 1)))

                (let ([sheet_list (XLSX-sheet_list (*XLSX*))])
                  (check-equal? (length sheet_list) 6)
                  (check-true (DATA-SHEET? (list-ref sheet_list 0)))
                  (check-true (DATA-SHEET? (list-ref sheet_list 1)))
                  (check-true (DATA-SHEET? (list-ref sheet_list 2)))
                  (check-true (CHART-SHEET? (list-ref sheet_list 3)))
                  (check-true (CHART-SHEET? (list-ref sheet_list 4)))
                  (check-true (CHART-SHEET? (list-ref sheet_list 5)))
                  )))
             )
           (lambda ()
             (when (file-exists? content_type_file) (delete-file content_type_file))))
       )))
   ))

(run-tests test-content-type)
