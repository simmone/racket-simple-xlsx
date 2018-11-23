#lang racket

(require rackunit/text-ui)

(require rackunit "../../../../../writer/xl/drawings/_rels/drawing-rels.rkt")

(define test-drawing-rels
  (test-suite
   "test-drawing-rels"

   (test-case
    "test-drawing-rels"

    (check-equal? (write-drawing-rels 2)
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart2.xml\"/></Relationships>")))
    ))

(run-tests test-drawing-rels)
