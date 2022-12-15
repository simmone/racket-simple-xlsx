#lang racket

(require "lib.rkt")

(provide (contract-out
          [struct BORDER-STYLE
                  (
                   (hash_code string?)
                   (top_color (or/c "" rgb?))
                   (top_mode border-mode?)
                   (bottom_color (or/c "" rgb?))
                   (bottom_mode border-mode?)
                   (left_color (or/c "" rgb?))
                   (left_mode border-mode?)
                   (right_color (or/c "" rgb?))
                   (right_mode border-mode?)
                   )]
          [border-mode? (-> string? boolean?)]
          [border-style<? (-> (or/c #f BORDER-STYLE?) (or/c #f BORDER-STYLE?) boolean?)]
          [border-style=? (-> (or/c #f BORDER-STYLE?) (or/c #f BORDER-STYLE?) boolean?)]
          [border-style-from-hash-code (-> string? (or/c #f BORDER-STYLE?))]
          [*BORDER_STYLE->INDEX_MAP* (parameter/c (or/c (hash/c string? natural?) #f))]
          [*BORDER_INDEX->STYLE_MAP* (parameter/c (or/c (hash/c natural? string?) #f))]
          ))

(define *BORDER_STYLE->INDEX_MAP* (make-parameter #f))
(define *BORDER_INDEX->STYLE_MAP* (make-parameter #f))

(struct BORDER-STYLE (hash_code top_color top_mode bottom_color bottom_mode left_color left_mode right_color right_mode)
        #:transparent
        #:guard
        (lambda (_hash_code _top_color _top_mode _bottom_color _bottom_mode _left_color _left_mode _right_color _right_mode name)
          (values
           (format "~a<p>~a<p>~a<p>~a<p>~a<p>~a<p>~a<p>~a"
                   (string-upcase _top_color) _top_mode (string-upcase _bottom_color) _bottom_mode (string-upcase _left_color) _left_mode (string-upcase _right_color) _right_mode)
           (string-upcase _top_color) _top_mode (string-upcase _bottom_color) _bottom_mode (string-upcase _left_color) _left_mode (string-upcase _right_color) _right_mode)))

(define (border-style-from-hash-code hash_code)
  (let ([items (regexp-split #rx"<p>" hash_code)])
    (if (= (length items) 8)
        (BORDER-STYLE
         ""
         (list-ref items 0)
         (list-ref items 1)
         (list-ref items 2)
         (list-ref items 3)
         (list-ref items 4)
         (list-ref items 5)
         (list-ref items 6)
         (list-ref items 7))
        #f)))

(define (border-mode? mode)
  (ormap (lambda (_mode) (string=? _mode mode)) '("" "thin" "dashed" "double" "thick" "medium")))

(define (border-style=? border1 border2)
  (cond
   [(and (equal? border1 #f) (equal? border2 #f))
    #t]
   [(or (equal? border1 #f) (equal? border2 #f))
    #f]
   [else
    (string=? (BORDER-STYLE-hash_code border1) (BORDER-STYLE-hash_code border2))]))

(define (border-style<? border1 border2)
  (cond
   [(and (equal? border1 #f) (equal? border2 #f))
    #f]
   [(equal? border1 #f)
    #t]
   [(equal? border2 #f)
    #f]
   [(not (string=? (BORDER-STYLE-top_color border1) (BORDER-STYLE-top_color border2)))
         (string<? (BORDER-STYLE-top_color border1) (BORDER-STYLE-top_color border2))]
   [(not (string=? (BORDER-STYLE-top_mode border1) (BORDER-STYLE-top_mode border2)))
         (string<? (BORDER-STYLE-top_mode border1) (BORDER-STYLE-top_mode border2))]
   [(not (string=? (BORDER-STYLE-bottom_color border1) (BORDER-STYLE-bottom_color border2)))
         (string<? (BORDER-STYLE-bottom_color border1) (BORDER-STYLE-bottom_color border2))]
   [(not (string=? (BORDER-STYLE-bottom_mode border1) (BORDER-STYLE-bottom_mode border2)))
         (string<? (BORDER-STYLE-bottom_mode border1) (BORDER-STYLE-bottom_mode border2))]
   [(not (string=? (BORDER-STYLE-left_color border1) (BORDER-STYLE-left_color border2)))
         (string<? (BORDER-STYLE-left_color border1) (BORDER-STYLE-left_color border2))]
   [(not (string=? (BORDER-STYLE-left_mode border1) (BORDER-STYLE-left_mode border2)))
         (string<? (BORDER-STYLE-left_mode border1) (BORDER-STYLE-left_mode border2))]
   [(not (string=? (BORDER-STYLE-right_color border1) (BORDER-STYLE-right_color border2)))
         (string<? (BORDER-STYLE-right_color border1) (BORDER-STYLE-right_color border2))]
   [(not (string=? (BORDER-STYLE-right_mode border1) (BORDER-STYLE-right_mode border2)))
         (string<? (BORDER-STYLE-right_mode border1) (BORDER-STYLE-right_mode border2))]
   [else
    #f]))
