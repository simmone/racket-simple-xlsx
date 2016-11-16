#lang racket

(require rackunit/text-ui)

(require rackunit "chartsheet-rels.rkt")

(define test-chart-sheet-rels
  (test-suite
   "test-chart-sheet-rels"

   (test-case
    "test-chart-sheet-rels"

    (check-equal? (write-chart-sheet-rels 2)
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing2.xml\"/></Relationships>")))
    ))

(run-tests test-chart-sheet-rels)
