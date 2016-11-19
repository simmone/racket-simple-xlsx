#lang racket

(provide (contract-out
          [struct xlsxDef
                  (
                   (sheets list?)
                   )]
          [struct sheetDef 
                  (
                   (name string?)
                   (seq exact-nonnegative-integer?)
                   (type symbol?)
                   (typeSeq exact-nonnegative-integer?)
                   (data any)
                   )]
          [struct sheetData
                  (
                   (data_list list?)
                   (col_attr_hash hash?)
                   )]
          [struct colAttr
                  (
                   (width exact-nonnegative-integer?)
                   (back_color string?)
                   )]
          [struct lineChartData
                  (
                   (topic string?)
                   (x_data_def string?)
                   (y_data_hash hash?)
                   )]
          ))

(struct xlsxDef ([data #:mutable]))

(struct sheetDef ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [data #:mutable]))

(struct sheetData ([data_list #:mutable] [col_attr_hash #:mutable]))

(struct lineChartData ([topic #:mutable] [x_data #:mutable] [y_data_hash #:mutable])

(struct colAttr ([width #:mutable] [back_color #:mutable]))
