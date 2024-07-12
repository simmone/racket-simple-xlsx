#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/drawings/drawing.rkt"
         racket/runtime-path)

(define-runtime-path drawing1_file "drawing1.xml")
(define-runtime-path drawing2_file "drawing2.xml")
(define-runtime-path drawing3_file "drawing3.xml")

(define test-drawing
  (test-suite
   "test-drawing"

   (test-case
    "test-write-drawings"

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
             (write-drawings (apply build-path (drop-right (explode-path drawing1_file) 1))))
           (lambda ()
             (call-with-input-file drawing1_file
               (lambda (expected1)
                 (call-with-input-file drawing2_file
                   (lambda (expected2)
                     (call-with-input-file drawing3_file
                       (lambda (expected3)
                         (call-with-input-string
                          (lists-to-xml (drawing 1))
                          (lambda (actual)
                            (check-lines? expected1 actual)))
                         (call-with-input-string
                          (lists-to-xml (drawing 2))
                          (lambda (actual)
                            (check-lines? expected2 actual)))
                         (call-with-input-string
                          (lists-to-xml (drawing 3))
                          (lambda (actual)
                            (check-lines? expected3 actual))))))))))
           (lambda ()
             (when (file-exists? drawing1_file) (delete-file drawing1_file))
             (when (file-exists? drawing2_file) (delete-file drawing2_file))
             (when (file-exists? drawing3_file) (delete-file drawing3_file))
             )))))
   ))

(run-tests test-drawing)
