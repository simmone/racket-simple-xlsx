#lang racket

(provide (contract-out
          [struct NUMBER-STYLE
                  (
                   (formatId (or/c #f string?))
                   (formatCode (or/c string? 'APP))
                   )]
          [update-number-style (-> NUMBER-STYLE? NUMBER-STYLE? void?)]
          ))

(struct NUMBER-STYLE
        (
         (formatId #:mutable)
         (formatCode #:mutable)
         )
        #:transparent
        )

(define (update-number-style number_style new_style)
  (set-NUMBER-STYLE-formatCode! number_style (NUMBER-STYLE-formatCode new_style)))
