#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../lib/lib.rkt")
(require "../../../lib/sheet-lib.rkt")

(require"../../../xl/sharedStrings.rkt")

(require racket/runtime-path)
(define-runtime-path shared_strings_file "sharedStrings.xml")

(define test-shared-strings
  (test-suite
   "test-shared-strings"

   (test-case
    "test-write-shared-strings"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("chenxiao" "love" "陈思衡")))

       (squash-shared-strings-map)

       (dynamic-wind
           (lambda ()
             (write-shared-strings (apply build-path (drop-right (explode-path shared_strings_file) 1))))
           (lambda ()
             (call-with-input-file shared_strings_file
               (lambda (expected)
                 (call-with-input-string
                  (lists->xml (to-shared-strings))
                  (lambda (actual)
                    (check-lines? expected actual)))))

             (hash-clear! (XLSX-shared_string->index_map (*XLSX*)))
             (hash-clear! (XLSX-shared_index->string_map (*XLSX*)))

             (check-equal? (hash-count (XLSX-shared_string->index_map (*XLSX*))) 0)
             (check-equal? (hash-count (XLSX-shared_index->string_map (*XLSX*))) 0)

             (read-shared-strings (apply build-path (drop-right (explode-path shared_strings_file) 1)))

             (check-equal? (hash-count (XLSX-shared_string->index_map (*XLSX*))) 3)
             (check-equal? (hash-count (XLSX-shared_index->string_map (*XLSX*))) 3)

             (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "chenxiao") 0)
             (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 0) "chenxiao")

             (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "love") 1)
             (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 1) "love")

             (check-equal? (hash-ref (XLSX-shared_string->index_map (*XLSX*)) "陈思衡") 2)
             (check-equal? (hash-ref (XLSX-shared_index->string_map (*XLSX*)) 2) "陈思衡")
             )
           (lambda ()
             (when (file-exists? shared_strings_file) (delete-file shared_strings_file))
             )))))
   ))

(run-tests test-shared-strings)
