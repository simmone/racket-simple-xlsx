#lang racket

(provide (contract-out
          [struct sheet 
                  ((name string?)
                   (r_id string?)
                   (type symbol?)
                   (data list?)
                   (col_attr_hash hash?)
                   )]
          ))

(struct sheet ([name #:mutable] [r_id #:mutable] [type #:mutable] [data #:mutable]))

(struct col_attr ([width #:mutable] [back_color #:mutable]))
