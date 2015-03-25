#lang racket

(require rackunit/text-ui)

(require rackunit "sharedStrings.rkt")

(define test-shared-strings
  (test-suite
   "test-shared-strings"

   (test-case
    "test-shared-strings"
    (check-equal? (write-shared-strings '("chenxiao" "love" "陈思衡"))
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\"><si><t>chenxiao</t><phoneticPr fontId=\"1\" type=\"noConversion\"/></si><si><t>love</t><phoneticPr fontId=\"1\" type=\"noConversion\"/></si><si><t>陈思衡</t><phoneticPr fontId=\"1\" type=\"noConversion\"/></si></sst>")))
   ))

(run-tests test-shared-strings)
