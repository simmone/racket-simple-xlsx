#lang racket

(require "../../style/styles.rkt")
(require "../../style/fill-style.rkt")

(provide (contract-out
          [to-fills (-> (listof FILL-STYLE?) list?)]
          [from-fills (-> hash? void?)]
          ))

(define (to-fills fill_list)
  (append
   (list "fills" (cons "count" (number->string (length fill_list))))
   (let loop ([fills fill_list]
              [result_list '()])
     (if (not (null? fills))
         (loop
          (cdr fills)
          (cons
           (list "fill"
                 (list
                  "patternFill"
                  (cons "patternType" (FILL-STYLE-pattern (car fills)))
                  (list "fgColor" (cons "rgb" (FILL-STYLE-color (car fills))))
                  '("bgColor" ("indexed" . "64"))))
           result_list))
         (reverse result_list)))))

(define (from-fills xml_hash)
  (let loop ([loop_count 0]
             [fill_list '()])
    (if (< loop_count (hash-ref xml_hash "styleSheet1.fills1.fill's count" 0))
        (let ([prefix (format "styleSheet1.fills1.fill~a" (add1 loop_count))])
          (let ([fill_style
                 (FILL-STYLE
                  (hash-ref xml_hash (format "~a.patternFill1.fgColor1.rgb" prefix) "FFFFFF")
                  (hash-ref xml_hash (format "~a.patternFill1.patternType" prefix) "none"))])
            (loop (add1 loop_count) (cons fill_style fill_list))))
        (set-STYLES-fill_list! (*STYLES*) (reverse fill_list)))))
