#lang racket

(require "lib.rkt")

(provide (contract-out
          [struct NUMBER-STYLE
                  (
                   (hash_code string?)
                   (format string?)
                   )]
          [number-style-from-hash-code (-> string? (or/c #f NUMBER-STYLE?))]
          [number-style<? (-> (or/c #f NUMBER-STYLE?) (or/c #f NUMBER-STYLE?) boolean?)]
          [number-style=? (-> (or/c #f NUMBER-STYLE?) (or/c #f NUMBER-STYLE?) boolean?)]
          [*NUMBER_STYLE->INDEX_MAP* (parameter/c (or/c (hash/c string? natural?) #f))]
          [*NUMBER_INDEX->STYLE_MAP* (parameter/c (or/c (hash/c exact-integer? string?) #f))]
          ))

(define *NUMBER_STYLE->INDEX_MAP* (make-parameter #f))
(define *NUMBER_INDEX->STYLE_MAP* (make-parameter #f))

(struct NUMBER-STYLE (hash_code format)
        #:guard
        (lambda (_hash_code _format name)
          (values (format "~a" _format) _format)))

(define (number-style-from-hash-code hash_code)
  (if (string=? hash_code "")
      #f
      (NUMBER-STYLE "" hash_code)))

(define (number-style=? number1 number2)
  (cond
   [(and (equal? number1 #f) (equal? number2 #f))
    #t]
   [(or (equal? number1 #f) (equal? number2 #f))
    #f]
   [else
    (string=? (NUMBER-STYLE-hash_code number1) (NUMBER-STYLE-hash_code number2))]))

(define (number-style<? number1 number2)
  (cond
   [(and (equal? number1 #f) (equal? number2 #f))
    #f]
   [(equal? number1 #f)
    #t]
   [(equal? number2 #f)
    #f]
   [(not (string=? (NUMBER-STYLE-format number1) (NUMBER-STYLE-format number2)))
         (string<? (NUMBER-STYLE-format number1) (NUMBER-STYLE-format number2))]
   [else
    #f]))
