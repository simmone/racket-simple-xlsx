#lang racket

(require "style-lib.rkt")

(provide (contract-out
          [struct BORDER-STYLE
                  (
                   (left_color (or/c #f rgb?))
                   (left_mode (or/c #f border-mode?))
                   (right_color (or/c #f rgb?))
                   (right_mode (or/c #f border-mode?))
                   (top_color (or/c #f rgb?))
                   (top_mode (or/c #f border-mode?))
                   (bottom_color (or/c #f rgb?))
                   (bottom_mode (or/c #f border-mode?))
                   )]
          [border-mode? (-> string? boolean?)]
          [new-border-style (-> BORDER-STYLE?)]
          [update-border-style (-> BORDER-STYLE? BORDER-STYLE? void?)]
          ))

(struct BORDER-STYLE
        (
         (left_color #:mutable)
         (left_mode #:mutable)
         (right_color #:mutable)
         (right_mode #:mutable)
         (top_color #:mutable)
         (top_mode #:mutable)
         (bottom_color #:mutable)
         (bottom_mode #:mutable)
         )
        #:transparent
        )

(define (new-border-style)
  (BORDER-STYLE #f #f #f #f #f #f #f #f))

(define (update-border-style border_style new_style)
  (when (BORDER-STYLE-left_color new_style)
    (set-BORDER-STYLE-left_color! border_style (BORDER-STYLE-left_color new_style)))

  (when (BORDER-STYLE-left_mode new_style)
    (set-BORDER-STYLE-left_mode! border_style (BORDER-STYLE-left_mode new_style)))

  (when (BORDER-STYLE-right_color new_style)
    (set-BORDER-STYLE-right_color! border_style (BORDER-STYLE-right_color new_style)))

  (when (BORDER-STYLE-right_mode new_style)
    (set-BORDER-STYLE-right_mode! border_style (BORDER-STYLE-right_mode new_style)))

  (when (BORDER-STYLE-top_color new_style)
    (set-BORDER-STYLE-top_color! border_style (BORDER-STYLE-top_color new_style)))

  (when (BORDER-STYLE-top_mode new_style)
    (set-BORDER-STYLE-top_mode! border_style (BORDER-STYLE-top_mode new_style)))

  (when (BORDER-STYLE-bottom_color new_style)
    (set-BORDER-STYLE-bottom_color! border_style (BORDER-STYLE-bottom_color new_style)))

  (when (BORDER-STYLE-bottom_mode new_style)
    (set-BORDER-STYLE-bottom_mode! border_style (BORDER-STYLE-bottom_mode new_style)))
  )

(define (border-mode? mode)
  (ormap (lambda (_mode)
           (string=? _mode mode))
         '("thin" "dashed" "double" "thick" "medium" "dotted" "mediumDashed" "none"
           "dashDot" "dashDotDot" "mediumDashDot" "mediumDashDotDot" "slantDashDot"
           "hair")))
