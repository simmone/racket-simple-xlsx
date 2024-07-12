#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../sheet/sheet.rkt"
         "../../../style/style.rkt"
         "../../../lib/lib.rkt"
         "../../../style/assemble-styles.rkt"
         "../../../style/set-styles.rkt"
         "../../../xl/styles/styles.rkt"
         racket/runtime-path)

(define-runtime-path styles_file "styles.xml")

(require "cellXfs/cellXfs-alignment-test.rkt")
(require "styles/styles-test.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-write-and-read-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-cellXfses)))

       (strip-styles)
       (assemble-styles)

       (dynamic-wind
           (lambda ()
             (write-styles (apply build-path (drop-right (explode-path styles_file) 1))))
           (lambda ()
             (call-with-input-file styles_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (to-styles))
                  (lambda (actual)
                    (check-lines? expected actual)))))

             (with-xlsx
              (lambda ()
                (from-styles styles_file)
                (check-styles))))
           (lambda ()
             (when (file-exists? styles_file) (delete-file styles_file)))))))
   ))

(run-tests test-styles)
