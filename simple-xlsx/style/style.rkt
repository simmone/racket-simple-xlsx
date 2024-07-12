#lang racket

(require "border-style.rkt"
         "font-style.rkt"
         "alignment-style.rkt"
         "number-style.rkt"
         "fill-style.rkt")

(provide (contract-out
          [struct STYLE
                  (
                   (border_style (or/c BORDER-STYLE? #f))
                   (font_style (or/c  FONT-STYLE? #f))
                   (alignment_style (or/c ALIGNMENT-STYLE? #f))
                   (number_style (or/c NUMBER-STYLE? #f))
                   (fill_style (or/c FILL-STYLE? #f))
                   )]
          [new-style (-> STYLE?)]
          ))

(struct STYLE (
               (border_style #:mutable)
               (font_style #:mutable)
               (alignment_style #:mutable)
               (number_style #:mutable)
               (fill_style #:mutable)
               )
        #:transparent
        )

(define (new-style)
  (STYLE #f #f #f #f #f))
