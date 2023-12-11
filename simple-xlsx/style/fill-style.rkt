#lang racket

(require "style-lib.rkt")

(provide (contract-out
          [struct FILL-STYLE
                  (
                   (color rgb?)
                   (pattern fill-pattern?)
                   )]
          [fill-pattern? (-> string? boolean?)]
          [update-fill-style (-> FILL-STYLE? FILL-STYLE? void?)]
          ))

(struct FILL-STYLE
        (
         (color #:mutable)
         (pattern #:mutable)
         )
        #:transparent
        )

(define (update-fill-style fill_style new_style)
  (set-FILL-STYLE-color! fill_style (FILL-STYLE-color new_style))
  (set-FILL-STYLE-pattern! fill_style (FILL-STYLE-pattern new_style))
  )

(define (fill-pattern? pattern)
  (ormap (lambda (_pattern) (string=? _pattern pattern))
         '("none"
           "solid" "gray125" "darkGray" "mediumGray" "lightGray"
           "gray0625" "darkHorizontal" "darkVertical" "darkDown" "darkUp"
           "darkGrid" "darkTrellis" "lightHorizontal" "lightVertical" "lightDown"
           "lightUp" "lightGrid" "lightTrellis")))
