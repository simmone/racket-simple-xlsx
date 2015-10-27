#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

;; strings list convert to (string . place) hash
(provide (contract-out
          [write-shared-strings (-> list? string?)]
          [write-shared-strings-file (-> path-string? list? void?)]
          [filter-string (-> string? string?)]
          ))

(define S string-append)

(define (write-shared-strings string_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="@|(number->string (length string_list))|" uniqueCount="@|(number->string (length string_list))|">@|(with-output-to-string
    (lambda () 
      (let loop ([strings string_list])
        (when (not (null? strings))
          (printf "<si><t>~a</t><phoneticPr fontId=\"1\" type=\"noConversion\"/></si>" (filter-string (car strings)))
          (loop (cdr strings))))))|</sst>
})

(define (filter-string str)
  (regexp-replace* 
   #rx"<"
   (regexp-replace*
    #rx">"
    str
    "\\&gt;")
   "\\&lt;"))

(define (write-shared-strings-file dir string_list)
  (with-output-to-file (build-path dir "sharedStrings.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-shared-strings string_list)))))
