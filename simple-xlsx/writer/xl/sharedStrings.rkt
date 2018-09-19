#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/file)
(require racket/list)
(require racket/contract)

(require "../../xlsx/xlsx.rkt")

;; strings list convert to (string . place) hash
(provide (contract-out
          [write-shared-strings (-> list? string?)]
          [write-shared-strings-file (-> path-string? (is-a?/c xlsx%) void?)]
          [filter-string (-> string? string?)]
          ))

(define S string-append)

(define (write-shared-strings string_item_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="@|(number->string (length string_item_list))|" uniqueCount="@|(number->string (length string_item_list))|">@|(with-output-to-string
    (lambda () 
      (let loop ([strings string_item_list])
        (when (not (null? strings))
          (printf "<si><t>~a</t><phoneticPr fontId=\"1\" type=\"noConversion\"/></si>" (filter-string (car strings)))
          (loop (cdr strings))))))|</sst>
})

(define (filter-string str)
  (regexp-replace*
   #rx"<"
   (regexp-replace* 
    #rx">"
    (regexp-replace*
     #rx"&amp;"
     "&amp;"
     str)
     "\\&gt;")
    "\\&lt;")
   )

(define (write-shared-strings-file dir xlsx)
  (when (> (hash-count (get-field string_item_map xlsx)) 0)
        (make-directory* dir)

        (with-output-to-file (build-path dir "sharedStrings.xml")
          #:exists 'replace
          (lambda ()
            (printf "~a" (write-shared-strings (send xlsx get-string-item-list)))))))
