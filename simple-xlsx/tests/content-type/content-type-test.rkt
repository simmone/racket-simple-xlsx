#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt"
         "../../lib/lib.rkt"
         "../../lib/sheet-lib.rkt"
         "../../content-type.rkt"
         rackunit/text-ui
         rackunit
         racket/runtime-path)

(define-runtime-path content_type_test1_file "content_type_test1.xml")
(define-runtime-path content_type_test2_file "content_type_test2.xml")
(define-runtime-path content_type_test3_file "content_type_test3.xml")
(define-runtime-path content_type_test4_file "content_type_test4.xml")

(define test-content-type
  (test-suite
   "test-content-type"

   (test-case
    "test-content-type1"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (call-with-input-file content_type_test1_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content (to-content-type))
            (lambda (actual)
              (check-lines? expected actual)))))))

    (with-xlsx
     (lambda ()
       (from-content-type content_type_test1_file)

       (let ([sheet_list (XLSX-sheet_list (*XLSX*))])
         (check-equal? (length sheet_list) 3)
         (check-true (DATA-SHEET? (list-ref sheet_list 0)))
         (check-true (DATA-SHEET? (list-ref sheet_list 1)))
         (check-true (DATA-SHEET? (list-ref sheet_list 2))))))
    )

   (test-case
    "test-content-type2"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("1")))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (squash-shared-strings-map)

       (call-with-input-file content_type_test2_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content (to-content-type))
            (lambda (actual)
              (check-lines? expected actual)))))))

    (with-xlsx
     (lambda ()
       (from-content-type content_type_test2_file)

       (let ([sheet_list (XLSX-sheet_list (*XLSX*))])
         (check-equal? (length sheet_list) 3)
         (check-true (DATA-SHEET? (list-ref sheet_list 0)))
         (check-true (DATA-SHEET? (list-ref sheet_list 1)))
         (check-true (DATA-SHEET? (list-ref sheet_list 2))))))
    )

   (test-case
    "test-content-type3"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (call-with-input-file content_type_test3_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content (to-content-type))
            (lambda (actual)
              (check-lines? expected actual)))))))

    (with-xlsx
     (lambda ()
       (from-content-type content_type_test3_file)

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

       (squash-shared-strings-map)

       (call-with-input-file content_type_test4_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content (to-content-type))
            (lambda (actual)
              (check-lines? expected actual)))))))
    (with-xlsx
     (lambda ()
       (from-content-type content_type_test3_file)

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

   ))

(run-tests test-content-type)
