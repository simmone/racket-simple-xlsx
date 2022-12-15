#lang racket

(require "lib.rkt")
(require "../lib/dimension.rkt")

(provide (contract-out
          [struct MERGE-STYLE
                  (
                   (hash_code string?)
                   (cell_range cell-range?)
                   )]
          [merge-style-from-hash-code (-> string? (or/c #f MERGE-STYLE?))]
          [merge-style<? (-> (or/c #f MERGE-STYLE?) (or/c #f MERGE-STYLE?) boolean?)]
          [merge-style=? (-> (or/c #f MERGE-STYLE?) (or/c #f MERGE-STYLE?) boolean?)]
          ))

(struct MERGE-STYLE (hash_code cell_range)
        #:guard
        (lambda (_hash_code _cell_range name)
          (values (format "~a" _cell_range) _cell_range)))

(define (merge-style-from-hash-code hash_code)
  (if (string=? hash_code "")
      #f
      (MERGE-STYLE "" hash_code)))

(define (merge-style=? merge1 merge2)
  (cond
   [(and (equal? merge1 #f) (equal? merge2 #f))
    #t]
   [(or (equal? merge1 #f) (equal? merge2 #f))
    #f]
   [else
    (string=? (MERGE-STYLE-hash_code merge1) (MERGE-STYLE-hash_code merge2))]))

(define (merge-style<? merge1 merge2)
  (cond
   [(and (equal? merge1 #f) (equal? merge2 #f))
    #f]
   [(equal? merge1 #f)
    #t]
   [(equal? merge2 #f)
    #f]
   [(not (string=? (MERGE-STYLE-cell_range merge1) (MERGE-STYLE-cell_range merge2)))
         (string<? (MERGE-STYLE-cell_range merge1) (MERGE-STYLE-cell_range merge2))]
   [else
    #f]))
