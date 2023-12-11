#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../style/style.rkt")
(require "../../../../style/assemble-styles.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/styles/styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "empty.xml")

(define test-empty-styles
  (test-suite
   "test-empty-styles"

   (test-case
    "test-empty-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (strip-styles)
       (assemble-styles)

       (call-with-input-file test_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content (to-styles))
            (lambda (actual)
              (check-lines? expected actual))))))))
     ))

(run-tests test-empty-styles)
