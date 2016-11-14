#lang racket

(provide (contract-out
          [struct sheet-data 
                  (
                   (name string?)
                   (seq exact-nonnegative-integer?)
                   (type symbol?)
                   (data list?)
                   (col_attr_hash hash?)
                   )]
          [struct col-attr
                  (
                   (width exact-nonnegative-integer?)
                   (back_color string?)
                   )]
          ))

(struct sheet-data ([name #:mutable] [seq #:mutable] [type #:mutable] [data #:mutable] [col_attr_hash #:mutable]))

(struct col-attr ([width #:mutable] [back_color #:mutable]))
