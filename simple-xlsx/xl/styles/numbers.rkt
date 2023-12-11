#lang racket

(require "../../style/styles.rkt")
(require "../../style/number-style.rkt")

(provide (contract-out
          [to-numbers (-> (listof NUMBER-STYLE?) list?)]
          [from-numbers (-> hash? void?)]
          ))

(define (to-numbers number_list)
  (append
   (list "numFmts" (cons "count" (number->string (length number_list))))
   (let loop ([numbers number_list]
              [result_list '()])
     (if (not (null? numbers))
         (loop
          (cdr numbers)
          (cons
           (list "numFmt"
                 (cons "numFmtId" (NUMBER-STYLE-formatId (car numbers)))
                 (cons "formatCode" (NUMBER-STYLE-formatCode (car numbers))))
           result_list))
         (reverse result_list)))))

(define (from-numbers xml_hash)
  (let loop ([loop_count 0]
             [number_list '()])
    (if (< loop_count (hash-ref xml_hash "styleSheet1.numFmts1.numFmt's count" 0))
        (let ([prefix (format "styleSheet1.numFmts1.numFmt~a" (add1 loop_count))])
          (when (hash-has-key? xml_hash (format "~a.formatCode" prefix))
            (let ([number_style
                   (NUMBER-STYLE
                    (hash-ref xml_hash (format "~a.numFmtId" prefix))
                    (hash-ref xml_hash (format "~a.formatCode" prefix)))])
              (loop (add1 loop_count) (cons number_style number_list)))))
        (set-STYLES-number_list! (*STYLES*) (reverse number_list)))))
