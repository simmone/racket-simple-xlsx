#lang racket

(provide (contract-out
          [rgb? (-> string? boolean?)]
          [equal-hash-code=? (-> any/c any/c boolean?)]
          ))

(define (equal-hash-code=? s1 s2)
  (= (equal-hash-code s1) (equal-hash-code s2)))

(define (rgb? color_string)
  (if (or
       (regexp-match #px"^([0-9]|[A-Z]){6}$" color_string)
       (regexp-match #px"^([0-9]|[A-Z]){8}$" color_string))
      #t
      #f))
