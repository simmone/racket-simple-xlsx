#lang racket

(provide (contract-out
          [load-shared-strings (-> path-string? hash?)]
          ))

(require simple-xml)

(define (load-shared-strings shared_string_file)
  (let ([xml_hash (xml->hash shared_string_file)]
        [shared_hash (make-hash)])

    (let loop ([loop_count 1])
      (when (<= loop_count (hash-ref xml_hash "sst.si's count" 0))
            (let ([t (hash-ref xml_hash (format "sst.si~a.t" loop_count))])
              (hash-set! shared_hash
                         loop_count
                         (cond
                          [(string? t)
                           t]
                          [(integer? t)
                           (string (integer->char t))]
                          [else
                           ""])))
            (loop (add1 loop_count))))
    shared_hash))
