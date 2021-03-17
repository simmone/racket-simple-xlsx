#lang racket

(provide (contract-out
          [struct DATA-SHEET
                  (
                   (name string?)
                   (rows list?)
                   (width_hash hash?)
                   (height_hash hash?)
                   (freeze_range (cons/c natural? natural?))
                   (cell_to_origin_style_hash hash?)
                   (cell_to_style_index_hash hash?)
                   (row_to_origin_style_hash hash?)
                   (row_to_style_index_hash hash?)
                   (col_to_origin_style_hash hash?)
                   (col_to_style_index_hash hash?)
                   )]
          [struct CHART-SHEET
                  (
                   (name string?)
                   (chart_type symbol?)
                   (topic string?)
                   (x_topic string?)
                   (x_data_range DATA-RANGE?)
                   (y_data_range_list list?)
                   )]
          [struct DATA-RANGE
                  (
                   (sheet_name string?)
                   (range_str string?)
                   )]
          [struct DATA-SERIAL
                  (
                   (topic string?)
                   (data_range DATA-RANGE?)
                   )]
          ))

(struct DATA-SHEET (
                    [name #:mutable]
                    [rows #:mutable] 
                    [width_hash #:mutable]
                    [height_hash #:mutable]
                    [freeze_range #:mutable]
                    [cell_to_origin_style_hash #:mutable]
                    [cell_to_style_index_hash #:mutable]
                    [row_to_origin_style_hash #:mutable]
                    [row_to_style_index_hash #:mutable]
                    [col_to_origin_style_hash #:mutable]
                    [col_to_style_index_hash #:mutable]
                    ))

(struct colAttr ([width #:mutable] [back_color #:mutable]))

(struct CHART-SHEET (
                     [name :mutable]
                     [chart_type #:mutable] 
                     [topic #:mutable] 
                     [x_topic #:mutable] 
                     [x_data_range #:mutable] 
                     [y_data_range_list #:mutable]))
(struct DATA-RANGE ([sheet_name #:mutable] [range_str #:mutable]))
(struct DATA-SERIAL ([topic #:mutable] [data_range #:mutable]))
