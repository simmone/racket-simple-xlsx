#lang racket

(require "style-lib.rkt")

(provide (contract-out
          [struct FONT-STYLE
                  (
                   (size number?)
                   (name string?)
                   (color rgb?)
                   )]
          [update-font-style (-> FONT-STYLE? FONT-STYLE? void?)]
          ))

(struct FONT-STYLE (
                    (size #:mutable)
                    (name #:mutable)
                    (color #:mutable)
                    )
        #:transparent
        )

(define (update-font-style font_style new_style)
  (set-FONT-STYLE-size! font_style (FONT-STYLE-size new_style))
  (set-FONT-STYLE-name! font_style (FONT-STYLE-name new_style))
  (set-FONT-STYLE-color! font_style (FONT-STYLE-color new_style)))
