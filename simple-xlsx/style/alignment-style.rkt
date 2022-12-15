#lang racket

(require "lib.rkt")

(provide (contract-out
          [struct ALIGNMENT-STYLE
                  (
                   (hash_code string?)
                   (horizontal_placement horizontal_mode?)
                   (vertical_placement vertical_mode?)
                   )]
          [horizontal_mode? (-> string? boolean?)]
          [vertical_mode? (-> string? boolean?)]
          [alignment-style-from-hash-code (-> string? (or/c #f ALIGNMENT-STYLE?))]
          [alignment-style<? (-> (or/c #f ALIGNMENT-STYLE?) (or/c #f ALIGNMENT-STYLE?) boolean?)]
          [alignment-style=? (-> (or/c #f ALIGNMENT-STYLE?) (or/c #f ALIGNMENT-STYLE?) boolean?)]
          ))

(define (horizontal_mode? mode)
  (ormap (lambda (_mode) (string=? _mode mode)) '("left" "right" "center")))

(define (vertical_mode? mode)
  (ormap (lambda (_mode) (string=? _mode mode)) '("top" "bottom" "center")))

(struct ALIGNMENT-STYLE (hash_code horizontal_placement vertical_placement)
        #:guard
        (lambda (_hash_code _horizontal_placement _vertical_placement name)
          (values
           (format "~a<p>~a" _horizontal_placement _vertical_placement)
           _horizontal_placement _vertical_placement)))

(define (alignment-style-from-hash-code hash_code)
  (let ([items (regexp-split #rx"<p>" hash_code)])
    (if (= (length items) 2)
        (ALIGNMENT-STYLE "" (first items) (second items))
        #f)))

(define (alignment-style=? alignment1 alignment2)
  (cond
   [(and (equal? alignment1 #f) (equal? alignment2 #f))
    #t]
   [(or (equal? alignment1 #f) (equal? alignment2 #f))
    #f]
   [else
    (string=? (ALIGNMENT-STYLE-hash_code alignment1) (ALIGNMENT-STYLE-hash_code alignment2))]))

(define (alignment-style<? alignment1 alignment2)
  (cond
   [(and (equal? alignment1 #f) (equal? alignment2 #f))
    #f]
   [(equal? alignment1 #f)
    #t]
   [(equal? alignment2 #f)
    #f]
   [(not (string=? (ALIGNMENT-STYLE-horizontal_placement alignment1) (ALIGNMENT-STYLE-horizontal_placement alignment2)))
         (string<? (ALIGNMENT-STYLE-horizontal_placement alignment1) (ALIGNMENT-STYLE-horizontal_placement alignment2))]
   [(not (string=? (ALIGNMENT-STYLE-vertical_placement alignment1) (ALIGNMENT-STYLE-vertical_placement alignment2)))
         (string<? (ALIGNMENT-STYLE-vertical_placement alignment1) (ALIGNMENT-STYLE-vertical_placement alignment2))]
   [else
    #f]))
