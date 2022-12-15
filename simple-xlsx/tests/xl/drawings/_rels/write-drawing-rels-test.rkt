#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../lib/lib.rkt")
(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")

(require"../../../../xl/drawings/_rels/drawing-rels.rkt")

(require racket/runtime-path)
(define-runtime-path drawing1_rels_file "drawing1.xml.rels")
(define-runtime-path drawing2_rels_file "drawing2.xml.rels")
(define-runtime-path drawing3_rels_file "drawing3.xml.rels")


(define test-drawing-rels
  (test-suite
   "test-drawing-rels"

   (test-case
    "test-write-drawing-rels"

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
            (write-drawings-rels (apply build-path (drop-right (explode-path drawing1_rels_file) 1))))
          (lambda ()
            (call-with-input-file drawing1_rels_file
              (lambda (expected1)
                (call-with-input-file drawing2_rels_file
                  (lambda (expected2)
                    (call-with-input-file drawing3_rels_file
                      (lambda (expected3)
                        (call-with-input-string
                         (lists->xml (drawing-rels 1))
                         (lambda (actual)
                           (check-lines? expected1 actual)))
                        (call-with-input-string
                         (lists->xml (drawing-rels 2))
                         (lambda (actual)
                           (check-lines? expected2 actual)))
                        (call-with-input-string
                         (lists->xml (drawing-rels 3))
                         (lambda (actual)
                           (check-lines? expected3 actual))))))))))
          (lambda ()
            (when (file-exists? drawing1_rels_file) (delete-file drawing1_rels_file))
            (when (file-exists? drawing2_rels_file) (delete-file drawing2_rels_file))
            (when (file-exists? drawing3_rels_file) (delete-file drawing3_rels_file))
            )))))
    ))

(run-tests test-drawing-rels)
