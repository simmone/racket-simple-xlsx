#lang racket

(provide (contract-out
          [struct sheet
                  (
                  (name string?)
                  (seq exact-nonnegative-integer?)
                  (type symbol?)
                  (typeSeq exact-nonnegative-integer?)
                  (content (or/c data-sheet? chart-sheet?))
                  )]
          [struct data-sheet
                  (
                   (rows list?)
                   (width_hash hash?)
                   (cell_to_origin_style_hash hash?)
                   (cell_to_style_index_hash hash?)
                   )]
          [struct chart-sheet
                  (
                   (chart_type symbol?)
                   (topic string?)
                   (x_topic string?)
                   (x_data_range data-range?)
                   (y_data_range_list list?)
                   )]
          [struct data-range
                  (
                   (sheet_name string?)
                   (range_str string?)
                   )]
          [struct data-serial
                  (
                   (topic string?)
                   (data_range data-range?)
                   )]
          ))

(struct sheet ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [content #:mutable]))

(struct data-sheet (
                    [rows #:mutable] 
                    [width_hash #:mutable]
                    [cell_to_origin_style_hash #:mutable]
                    [cell_to_style_index_hash #:mutable]
                    ))

(struct colAttr ([width #:mutable] [back_color #:mutable]))

(struct chart-sheet (
                     [chart_type #:mutable] 
                     [topic #:mutable] 
                     [x_topic #:mutable] 
                     [x_data_range #:mutable] 
                     [y_data_range_list #:mutable]))
(struct data-range ([sheet_name #:mutable] [range_str #:mutable]))
(struct data-serial ([topic #:mutable] [data_range #:mutable]))
