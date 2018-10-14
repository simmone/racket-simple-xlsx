#lang racket

(require "../../../xlsx/xlsx.rkt")

(require "../../../lib/lib.rkt")

(require rackunit/text-ui)

(require rackunit "../content-type.rkt")

(require racket/runtime-path)
(define-runtime-path test1_file "content-type-test1.dat")
(define-runtime-path test2_file "content-type-test2.dat")

(define test-content-type
  (test-suite
   "test-content-type"

   (test-case
    "test-content-type"
    
    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "Sheet1" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet2" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet3" #:sheet_data '((1)))
      (send xlsx add-chart-sheet #:sheet_name "Chart1" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart2" #:topic "Chart2" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart3" #:topic "Chart3" #:x_topic "")

      (call-with-input-file test1_file
        (lambda (expected)
          (call-with-input-string
           (write-content-type xlsx)
           (lambda (actual)
             (check-lines? expected actual))))))

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "Sheet1" #:sheet_data '(("1")))
      (send xlsx add-data-sheet #:sheet_name "Sheet2" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet3" #:sheet_data '((1)))
      (send xlsx add-chart-sheet #:sheet_name "Chart1" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart2" #:topic "Chart2" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart3" #:topic "Chart3" #:x_topic "")

      (call-with-input-file test2_file
        (lambda (expected)
          (call-with-input-string
           (write-content-type xlsx)
           (lambda (actual)
             (check-lines? expected actual))))))
    ))
   )

(run-tests test-content-type)
