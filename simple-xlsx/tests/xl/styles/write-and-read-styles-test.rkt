#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../style/style.rkt")
(require "../../../lib/lib.rkt")
(require "../../../style/assemble-styles.rkt")
(require "../../../style/set-styles.rkt")

(require"../../../xl/styles/styles.rkt")

(require racket/runtime-path)
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
                  (lists->xml (to-styles))
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
