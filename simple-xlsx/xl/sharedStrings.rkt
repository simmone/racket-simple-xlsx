#lang racket

(require simple-xml)

(require "../xlsx/xlsx.rkt")

(provide (contract-out
          [to-shared-strings (-> list?)]
          [from-shared-strings (-> path-string? void?)]
          [filter-string (-> string? string?)]
          [write-shared-strings (->* () (path-string?) void?)]
          [read-shared-strings (->* () (path-string?) void?)]
          ))

(define (to-shared-strings)
  (append
   '("sst" ("xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
   (list
    (cons "count" (format "~a" (hash-count (XLSX-shared_string->index_map (*XLSX*)))))
    (cons "uniqueCount" (format "~a" (hash-count (XLSX-shared_string->index_map (*XLSX*))))))
   (let loop ([strings (map car (sort (hash->list (XLSX-shared_string->index_map (*XLSX*))) < #:key cdr))]
              [result_list '()])
     (if (not (null? strings))
         (loop
          (cdr strings)
          (cons
           (list
            "si"
            (list "t" (filter-string (car strings)))
            (list "phoneticPr" (cons "fontId" "1") (cons "type" "noConversion")))
           result_list))
         (reverse result_list)))))

(define (from-shared-strings shared_strings_file)
  (when (file-exists? shared_strings_file)
        (let ([xml_hash (xml->hash shared_strings_file)])
          (let loop ([loop_count 0])
            (when (< loop_count (hash-ref xml_hash "sst1.si's count" 0))
                  (hash-set! (XLSX-shared_string->index_map (*XLSX*))
                             (hash-ref xml_hash (format "sst1.si~a.t1" (add1 loop_count)))
                             loop_count)
                  (hash-set! (XLSX-shared_index->string_map (*XLSX*))
                             loop_count
                             (hash-ref xml_hash (format "sst1.si~a.t1" (add1 loop_count))))
                  (loop (add1 loop_count)))))))

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

(define (write-shared-strings [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (make-directory* dir)

    (when (> (hash-count (XLSX-shared_string->index_map (*XLSX*))) 0)
        (with-output-to-file (build-path dir "sharedStrings.xml")
          #:exists 'replace
          (lambda ()
            (printf "~a" (lists->xml (to-shared-strings))))))))

(define (read-shared-strings [input_dir #f])
  (let ([dir (if input_dir input_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (from-shared-strings (build-path dir "sharedStrings.xml"))))
