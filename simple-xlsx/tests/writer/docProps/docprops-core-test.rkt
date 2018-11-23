#lang racket

(require rackunit/text-ui)

(require rackunit "../../../writer/docProps/docprops-core.rkt")

(require "../../../lib/lib.rkt")

(define test-docprops-app
  (test-suite
   "test-docprops-app"

   (test-case
    "test-docprops-app"
    (check-equal? (write-docprops-core (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST"))
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">2006-09-13T11:21:51Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2015-01-02T13:17:44+08:00</dcterms:modified></cp:coreProperties>")))

   ))

(run-tests test-docprops-app)
