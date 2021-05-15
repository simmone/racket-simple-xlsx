#lang racket

(require simple-xml)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../lib/lib.rkt")

(require rackunit/text-ui)

(require "../../../writer.rkt")

(require rackunit "../../../new/content-type.rkt")

(require racket/runtime-path)
(define-runtime-path test1_file "content-type-test1.xml")
(define-runtime-path test2_file "content-type-test2.xml")
(define-runtime-path test3_file "content-type-test3.xml")
(define-runtime-path test4_file "content-type-test4.xml")

(define test-content-type
  (test-suite
   "test-content-type"

   (test-case
    "test-content-type1"
    
    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (add-data-sheet "Sheet1" '((1)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))

      (call-with-input-file test1_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (content-type))
           (lambda (actual)
             (check-lines? expected actual)))))))

   (test-case
    "test-content-type2"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (add-data-sheet "Sheet1" '(("1")))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      
      (call-with-input-file test2_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (content-type))
           (lambda (actual)
             (check-lines? expected actual)))))))

   (test-case
    "test-content-type3"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (add-data-sheet "Sheet1" '(("1")))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      (add-chart-sheet "Chart1" 'LINE "Chart1")
      (add-chart-sheet "Chart2" 'LINE "Chart2")
      (add-chart-sheet "Chart3" 'LINE "Chart3")
      
      (call-with-input-file test3_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (content-type))
           (lambda (actual)
             (check-lines? expected actual)))))))

   (test-case
    "test-content-type4"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
      (add-data-sheet "Sheet1" '((1)))
      (add-data-sheet "Sheet2" '((1)))
      (add-data-sheet "Sheet3" '((1)))
      (add-chart-sheet "Chart1" 'LINE "Chart1")
      (add-chart-sheet "Chart2" 'LINE "Chart2")
      (add-chart-sheet "Chart3" 'LINE "Chart3")
      
      (call-with-input-file test4_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (content-type))
           (lambda (actual)
             (check-lines? expected actual)))))))
   ))

(run-tests test-content-type)
