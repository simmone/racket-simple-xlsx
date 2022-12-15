#lang racket

(require "lib.rkt")

(provide (contract-out
          [struct FONT-STYLE
                  (
                   (hash_code string?)
                   (size natural?)
                   (name string?)
                   (color rgb?)
                   )]
          [font-style-from-hash-code (-> string? (or/c #f FONT-STYLE?))]
          [font-style<? (-> (or/c #f FONT-STYLE?) (or/c #f FONT-STYLE?) boolean?)]
          [font-style=? (-> (or/c #f FONT-STYLE?) (or/c #f FONT-STYLE?) boolean?)]
          [*FONT_STYLE->INDEX_MAP* (parameter/c (or/c (hash/c string? natural?) #f))]
          [*FONT_INDEX->STYLE_MAP* (parameter/c (or/c (hash/c natural? string?) #f))]
          ))

(define *FONT_STYLE->INDEX_MAP* (make-parameter #f))
(define *FONT_INDEX->STYLE_MAP* (make-parameter #f))

(struct FONT-STYLE (hash_code size name color)
        #:guard
        (lambda (_hash_code _size _name _color name)
          (values
           (format "~a<p>~a<p>~a" _size _name (string-upcase _color))
           _size _name (string-upcase _color))))

(define (font-style-from-hash-code hash_code)
  (let ([items (regexp-split #rx"<p>" hash_code)])
    (if (= (length items) 3)
        (FONT-STYLE "" (string->number (first items)) (second items)  (third items))
        #f)))

(define (font-style=? font1 font2)
  (cond
   [(and (equal? font1 #f) (equal? font2 #f))
    #t]
   [(or (equal? font1 #f) (equal? font2 #f))
    #f]
   [else
    (string=? (FONT-STYLE-hash_code font1) (FONT-STYLE-hash_code font2))]))

(define (font-style<? font1 font2)
  (cond
   [(and (equal? font1 #f) (equal? font2 #f))
    #f]
   [(equal? font1 #f)
    #t]
   [(equal? font2 #f)
    #f]
   [(not (= (FONT-STYLE-size font1) (FONT-STYLE-size font2)))
         (< (FONT-STYLE-size font1) (FONT-STYLE-size font2))]
   [(not (string=? (FONT-STYLE-name font1) (FONT-STYLE-name font2)))
         (string<? (FONT-STYLE-name font1) (FONT-STYLE-name font2))]
   [(not (string=? (FONT-STYLE-color font1) (FONT-STYLE-color font2)))
         (string<? (FONT-STYLE-color font1) (FONT-STYLE-color font2))]
   [else
    #f]))
