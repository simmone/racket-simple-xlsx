#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../lib/lib.rkt")
(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../style/style.rkt")
(require "../../../../style/number-style.rkt")
(require "../../../../style/font-style.rkt")
(require "../../../../style/fill-style.rkt")
(require "../../../../style/border-style.rkt")
(require "../../../../style/sort-styles.rkt")
(require "../../../../style/set-styles.rkt")
(require "../../../../xl/styles/styles.rkt")

(require racket/runtime-path)
(define-runtime-path styles_file "styles_test.xml")
(define-runtime-path complex_styles_file "complex_styles_test.xml")

(require "../borders-test.rkt")
(require "../fills-test.rkt")
(require "../fonts-test.rkt")
(require "../numbers-test.rkt")
(require "../cellXfs-alignment-test.rkt")

(provide (contract-out
          [check-styles (-> void?)]
          ))

(define (check-styles)
  (check-border-styles)

  (check-fill-styles)

  (check-font-styles)

  (check-number-styles)

  (check-cellXfses))

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-to-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-cellXfses)))

       (sort-styles)

       (call-with-input-file styles_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content (to-styles))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-styles styles_file)

       (check-styles)
       )))

   (test-case
    "test-from-complex-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-styles complex_styles_file)
       )))

   ))

(run-tests test-styles)
