#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../xlsx/xlsx.rkt"
         "../../lib/lib.rkt"
         "../../_rels/rels.rkt"
         racket/runtime-path)

(define-runtime-path rels_file ".rels")

(define test-rels
  (test-suite
   "test-rels"

   (test-case
    "test-write-rels"

    (with-xlsx
     (lambda ()
       (dynamic-wind
           (lambda ()
             (write-rels (apply build-path (drop-right (explode-path rels_file) 1))))
           (lambda ()
             (call-with-input-file rels_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (rels))
                  (lambda (actual)
                    (check-lines? expected actual))))))
           (lambda ()
             (when (file-exists? rels_file) (delete-file rels_file)))))))
   ))

(run-tests test-rels)
