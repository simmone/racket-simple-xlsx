#lang racket

(provide (contract-out
          [struct ALIGNMENT-STYLE
                  (
                   (horizontal_placement horizontal_mode?)
                   (vertical_placement vertical_mode?)
                   )]
          [horizontal_mode? (-> string? boolean?)]
          [vertical_mode? (-> string? boolean?)]
          [update-alignment-style (-> ALIGNMENT-STYLE? ALIGNMENT-STYLE? void?)]
          ))

(define (horizontal_mode? mode)
  (ormap (lambda (_mode) (string=? _mode mode)) '("" "left" "right" "center" "general")))

(define (vertical_mode? mode)
  (ormap (lambda (_mode) (string=? _mode mode)) '("" "top" "bottom" "center" "general")))

(struct ALIGNMENT-STYLE
        (
         (horizontal_placement #:mutable)
         (vertical_placement #:mutable)
         )
        #:transparent
        )

(define (update-alignment-style border_style new_style)
  (when (not (string=? (ALIGNMENT-STYLE-horizontal_placement new_style) ""))
    (set-ALIGNMENT-STYLE-horizontal_placement! border_style (ALIGNMENT-STYLE-horizontal_placement new_style)))

  (when (not (string=? (ALIGNMENT-STYLE-vertical_placement new_style) ""))
    (set-ALIGNMENT-STYLE-vertical_placement! border_style (ALIGNMENT-STYLE-vertical_placement new_style))))
