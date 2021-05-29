#lang racket

(require simple-xml)

(require rackunit/text-ui)

(require rackunit "../../../../new/xl/sharedStrings.rkt")

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../writer.rkt")
(require "../../../../lib/lib.rkt")

(require racket/runtime-path)
(define-runtime-path shared_strings_file "sharedStrings.xml")

(define test-shared-strings
  (test-suite
   "test-shared-strings"

   (test-case
    "test-shared-strings"

    (parameterize 
     ([*CURRENT_XLSX* (new-xlsx)])
     (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))
     
     (call-with-input-file shared_strings_file
        (lambda (expected)
          (call-with-input-string
           (lists->xml (shared-strings))
           (lambda (actual)
             (check-lines? expected actual)))))))

   (test-case
    "test-filter-string"
    (check-equal? (filter-string "<a>") "&lt;a&gt;")
    (check-equal? (filter-string "<&a>") "&lt;&amp;a&gt;")
    (check-equal? (filter-string "<<a>>") "&lt;&lt;a&gt;&gt;")
   )))

(run-tests test-shared-strings)
