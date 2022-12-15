#lang racket

(require "../../style/fill-style.rkt")

(provide (contract-out
          [to-fills (-> (listof FILL-STYLE?) list?)]
          [from-fills (-> hash? void?)]
          ))

(define (to-fills fill_list)
  (append
   (list "fills" (cons "count" (number->string (+ 2 (length fill_list)))))
   '(("fill" ("patternFill" ("patternType" . "none"))))
   '(("fill" ("patternFill" ("patternType" . "gray125"))))
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
  (let ([skip_count 2])
    (let loop ([loop_count 0])
      (when (< loop_count (hash-ref xml_hash "styleSheet1.fills1.fill's count" 0))
            (when (>= loop_count skip_count)
                  (let ([prefix (format "styleSheet1.fills1.fill~a" (add1 loop_count))])
                    (let ([hash_code
                           (FILL-STYLE-hash_code
                            (FILL-STYLE
                             ""
                             (hash-ref xml_hash (format "~a.patternFill1.fgColor1.rgb" prefix) "000000")
                             (hash-ref xml_hash (format "~a.patternFill1.patternType" prefix) "none")))])

                      (hash-set! (*FILL_STYLE->INDEX_MAP*) hash_code (- loop_count skip_count))
                      (hash-set! (*FILL_INDEX->STYLE_MAP*) (- loop_count skip_count) hash_code))))
            (loop (add1 loop_count))))))

