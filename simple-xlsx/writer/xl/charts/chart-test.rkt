#lang racket

(require rackunit/text-ui)

(require rackunit "chartsheet.rkt")

(define test-chart
  (test-suite
   "test-chart"

   (test-case
    "test-chart"

    (check-equal? (write-chart 1)
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetPr/><sheetViews><sheetView zoomScale=\"115\" workbookViewId=\"0\" zoomToFit=\"1\"/></sheetViews><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><drawing r:id=\"rId1\"/></chartsheet>"))
    )))

(run-tests test-chart)
