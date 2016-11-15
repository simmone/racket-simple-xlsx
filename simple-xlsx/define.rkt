#lang racket

(provide (contract-out
          [struct sheetData 
                  (
                   (name string?)
                   (seq exact-nonnegative-integer?)
                   (type symbol?)
                   (typeSeq exact-nonnegative-integer?)
                   (data list?)
                   (col_attr_hash hash?)
                   )]
          [struct colAttr
                  (
                   (width exact-nonnegative-integer?)
                   (back_color string?)
                   )]
          ))

(struct sheetData ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [data #:mutable] [col_attr_hash #:mutable]))

(struct colAttr ([width #:mutable] [back_color #:mutable]))
