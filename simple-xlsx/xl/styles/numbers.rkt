#lang racket

(require "../../style/number-style.rkt")

(provide (contract-out
          [to-numbers (-> (listof NUMBER-STYLE?) list?)]
          [from-numbers (-> hash? void?)]
          ))

(define (to-numbers number_list)
  (append
   (list
    "numFmts" (cons "count" (format "~a" (add1 (length number_list)))))
   '(("numFmt" ("numFmtId" . "164") ("formatCode" . "General")))
   (let loop ([numbers number_list]
              [loop_num_id 164]
              [result_list '()])
     (if (not (null? numbers))
         (loop
          (cdr numbers)
          (add1 loop_num_id)
          (cons
           (list "numFmt" (cons "numFmtId" (format "~a" (add1 loop_num_id))) (cons "formatCode" (NUMBER-STYLE-format (car numbers))))
           result_list))
         (reverse result_list)))))

(define (from-numbers xml_hash)
  (let ([skip_count 1])
    (let loop ([loop_count 0])
      (when (< loop_count (hash-ref xml_hash "styleSheet1.numFmts1.numFmt's count" 0))
            (when (>= loop_count skip_count)
                  (let ([prefix (format "styleSheet1.numFmts1.numFmt~a" (add1 loop_count))])
                    (when (hash-has-key? xml_hash (format "~a.formatCode" prefix))
                          (let ([number_style_hash_code
                                 (NUMBER-STYLE-hash_code
                                  (NUMBER-STYLE
                                   ""
                                   (hash-ref xml_hash (format "~a.formatCode" prefix))))])

                            (hash-set! (*NUMBER_STYLE->INDEX_MAP*) number_style_hash_code (- loop_count skip_count))
                            (hash-set! (*NUMBER_INDEX->STYLE_MAP*) (- loop_count skip_count) number_style_hash_code)))))
            (loop (add1 loop_count))))))

