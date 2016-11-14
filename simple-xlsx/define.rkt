#lang racket

(provide (contract-out
          [struct sheet 
                  ((sheet_name string?)
                   (sheet_id string?)
                   (sheet_type symbol?)
                   (sheet_data list?)
                   (col_attr_hash hash?)
                   )]
          ))

(struct sheet (name r_id type data))

(struct col_attr (width back_color))
