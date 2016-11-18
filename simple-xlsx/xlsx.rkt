#lang racket

(provide (contract-out
          [xlsx% class?]
          ))

(struct sheet ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [data #:mutable]))

(struct sheetData ([data_list #:mutable] [col_attr_hash #:mutable]))

(struct lineChartData ([topic #:mutable] [x_data #:mutable] [y_data_hash #:mutable])

(struct colAttr ([width #:mutable] [back_color #:mutable]))

(define xlsx%
  (class object%
         (super-new)

         (field [sheet_list '()])))
