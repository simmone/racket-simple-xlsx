#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/file)
(require racket/list)
(require racket/contract)

(require simple-xml)

(require "../../xlsx/xlsx.rkt")

;; strings list convert to (string . place) hash
(provide (contract-out
          [write-shared-strings (-> list? string?)]
          [write-shared-strings-file (-> path-string? XLSX? void?)]
          [filter-string (-> string? string?)]
          [load-shared-strings (-> path-string? XLSX? void?)]
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

(define (write-shared-strings-file)
  (when (> (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*))) 0)
        (make-directory* dir)

        (with-output-to-file (build-path (XlSX-xlsx_dir (*CURRENT_XLSX*)) "xl" "sharedStrings.xml")
          #:exists 'replace
          (lambda ()
            (printf "~a" (write-shared-strings (send xlsx get-string-item-list)))))))

(define (load-shared-strings)
  (let ([xml_hash (xml->hash (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "xl" "sharedStrings.xml"))])
    (let loop ([loop_count 1])
      (when (<= loop_count (hash-ref xml_hash "sst.si's count" 0))
            (let ([t (hash-ref xml_hash (format "sst.si~a.t" loop_count))])
              (hash-set! (XLSX-shared_strings_map (*CURRENT_XLSX*))
                         (sub1 loop_count)
                         (cond
                          [(string? t)
                           t]
                          [(integer? t)
                           (string (integer->char t))]
                          [else
                           ""])))
            (loop (add1 loop_count))))))
