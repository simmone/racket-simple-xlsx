#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../lib/lib.rkt")
(require "../../../lib/sheet-lib.rkt")
(require"../../../xl/sharedStrings.rkt")

(require racket/runtime-path)
(define-runtime-path shared_strings_file "sharedStrings_test.xml")
(define-runtime-path shared_strings_not_exist_file "sharedStrings_not_exist.xml")

(define test-shared-strings
  (test-suite
   "test-shared-strings"

   (test-case
    "test-shared-strings"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))

       (squash-shared-strings-map)

       (check-equal? (hash-count (XLSX-shared_string->index_map (*XLSX*))) 3)
       (check-equal? (hash-count (XLSX-shared_index->string_map (*XLSX*))) 3)

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "chenxiao") 0)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 0) "chenxiao")

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "love") 1)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 1) "love")

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "陈思衡") 2)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 2) "陈思衡")

       (call-with-input-file shared_strings_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml (to-shared-strings))
            (lambda (actual)
              (check-lines? expected actual)))))))

    (with-xlsx
     (lambda ()
       (from-shared-strings shared_strings_file)

       (check-equal? (hash-count (XLSX-shared_string->index_map (*XLSX*))) 3)
       (check-equal? (hash-count (XLSX-shared_index->string_map (*XLSX*))) 3)

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "chenxiao") 0)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 0) "chenxiao")

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "love") 1)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 1) "love")

       (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "陈思衡") 2)
       (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 2) "陈思衡")
       ))

    (with-xlsx
     (lambda ()
       (from-shared-strings shared_strings_not_exist_file)

       (check-equal? (hash-count (XLSX-shared_string->index_map (*XLSX*))) 0)
       (check-equal? (hash-count (XLSX-shared_index->string_map (*XLSX*))) 0)
       ))

    )


   (test-case
    "test-filter-string"
    (check-equal? (filter-string "<a>") "&lt;a&gt;")
    (check-equal? (filter-string "<&a>") "&lt;&amp;a&gt;")
    (check-equal? (filter-string "<<a>>") "&lt;&lt;a&gt;&gt;"))

   ))

(run-tests test-shared-strings)
