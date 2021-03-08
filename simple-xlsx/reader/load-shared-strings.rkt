#lang racket

(provide (contract-out
          [load-shared-strings (-> path-string? XLSX? void?)]
          ))

(require simple-xml)

(require "../xlsx/xlsx.rkt")

(define (load-shared-strings shared_string_file _xlsx)
  (let ([xml_hash (xml->hash shared_string_file)])
    (let loop ([loop_count 1])
      (when (<= loop_count (hash-ref xml_hash "sst.si's count" 0))
            (let ([t (hash-ref xml_hash (format "sst.si~a.t" loop_count))])
              (hash-set! (XLSX-shared_strings_map _xlsx)
                         loop_count
                         (cond
                          [(string? t)
                           t]
                          [(integer? t)
                           (string (integer->char t))]
                          [else
                           ""])))
            (loop (add1 loop_count))))))

