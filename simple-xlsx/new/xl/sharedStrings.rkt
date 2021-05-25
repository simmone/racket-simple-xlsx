#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")

(provide (contract-out
          [shared-strings (-> list?)]
          [write-shared-strings-file (-> void?)]
          [filter-string (-> string? string?)]
          [read-shared-strings (-> void?)]
          ))

(define (shared-strings)
  (append
   ("sst" ("xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
   (list
    (cons "count" (format "~a" (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*)))))
    (cons "uniqueCount" (format "~a" (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*))))))
   (list
    (let loop ([strings (map car (sort (hash->list (XLSX->shared_strings_map (*CURRENT_XLSX*))) < #:key cdr))]
               [result_list '()])
      (if (not (null? strings))
          (loop
           (cdr strings)
           (cons
            (list
             "si"
             (list "t" (filter-string (car strings)))
             (list "phoneticPr" (cons "fontId" "1") (cons "type" "noConversion")))
            result_list)
           (reverse result_list)))))))

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
  (let ([dir (build-path (XlSX-xlsx_dir (*CURRENT_XLSX*)) "xl")])
    (make-directory* dir)

    (when (> (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*))) 0)
        (with-output-to-file (build-path dir "sharedStrings.xml")
          #:exists 'replace
          (lambda ()
            (printf "~a" (lists->compact_xml (shared-strings (send xlsx get-string-item-list)))))))))

(define (read-shared-strings)
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
