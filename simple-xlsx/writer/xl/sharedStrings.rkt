#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/file)
(require racket/list)
(require racket/contract)

(require "../../xlsx.rkt")

;; strings list convert to (string . place) hash
(provide (contract-out
          [write-shared-strings (-> hash? string?)]
          [write-shared-strings-file (-> path-string? hash? void?)]
          [filter-string (-> string? string?)]
          ))

(define S string-append)

(define (write-shared-strings string_item_map) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="@|(number->string (hash-count string_item_map))|" uniqueCount="@|(number->string (hash-count string_item_map))|">@|(with-output-to-string
    (lambda () 
      (let loop ([strings (sort (hash-keys string_item_map) string<?)])
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
  (make-directory* dir)

  (with-output-to-file (build-path dir "sharedStrings.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-shared-strings (get-field string_item_map xlsx))))))
