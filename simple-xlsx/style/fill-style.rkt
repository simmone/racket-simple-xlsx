#lang racket

(require "lib.rkt")

(provide (contract-out
          [struct FILL-STYLE
                  (
                   (hash_code string?)
                   (color rgb?)
                   (pattern fill-pattern?)
                   )]
          [fill-pattern? (-> string? boolean?)]
          [fill-style-from-hash-code (-> string? (or/c #f FILL-STYLE?))]
          [fill-style<? (-> (or/c #f FILL-STYLE?) (or/c #f FILL-STYLE?) boolean?)]
          [fill-style=? (-> (or/c #f FILL-STYLE?) (or/c #f FILL-STYLE?) boolean?)]
          [*FILL_STYLE->INDEX_MAP* (parameter/c (or/c (hash/c string? natural?) #f))]
          [*FILL_INDEX->STYLE_MAP* (parameter/c (or/c (hash/c natural? string?) #f))]
          ))

(define *FILL_STYLE->INDEX_MAP* (make-parameter #f))
(define *FILL_INDEX->STYLE_MAP* (make-parameter #f))

(struct FILL-STYLE (hash_code color pattern)
        #:guard
        (lambda (_hash_code _color _pattern name)
          (values
           (format "~a<p>~a" (string-upcase _color) _pattern)
           (string-upcase _color) _pattern)))

(define (fill-style-from-hash-code hash_code)
  (let ([items (regexp-split #rx"<p>" hash_code)])
    (if (= (length items) 2)
        (FILL-STYLE "" (first items) (second items))
        #f)))

(define (fill-pattern? pattern)
  (ormap (lambda (_pattern) (string=? _pattern pattern))
         '("none"
           "solid" "gray125" "darkGray" "mediumGray" "lightGray"
           "gray0625" "darkHorizontal" "darkVertical" "darkDown" "darkUp"
           "darkGrid" "darkTrellis" "lightHorizontal" "lightVertical" "lightDown"
           "lightUp" "lightGrid" "lightTrellis")))

(define (fill-style=? fill1 fill2)
  (cond
   [(and (equal? fill1 #f) (equal? fill2 #f))
    #t]
   [(or (equal? fill1 #f) (equal? fill2 #f))
    #f]
   [else
    (string=? (FILL-STYLE-hash_code fill1) (FILL-STYLE-hash_code fill2))]))

(define (fill-style<? fill1 fill2)
  (cond
   [(and (equal? fill1 #f) (equal? fill2 #f))
    #f]
   [(equal? fill1 #f)
    #t]
   [(equal? fill2 #f)
    #f]
   [(not (string=? (FILL-STYLE-color fill1) (FILL-STYLE-color fill2)))
         (string<? (FILL-STYLE-color fill1) (FILL-STYLE-color fill2))]
   [(not (string=? (FILL-STYLE-pattern fill1) (FILL-STYLE-pattern fill2)))
         (string<? (FILL-STYLE-pattern fill1) (FILL-STYLE-pattern fill2))]
   [else
    #f]))
