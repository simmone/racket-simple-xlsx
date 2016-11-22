#lang racket

(provide (contract-out
          [xlsx% class?]
          [struct sheet
                  (
                  (name string?)
                  (seq exact-nonnegative-integer?)
                  (type symbol?)
                  (typeSeq exact-nonnegative-integer?)
                  (content (or/c data-sheet? line-chart-sheet?))
                  )]
          [struct data-sheet
                  (
                   (rows list?)
                   (width_hash hash?)
                   (color_hash hash?)
                   )]
          [struct line-chart-sheet
                  (
                   (topic string?)
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

(struct data-sheet ([rows #:mutable] [width_hash #:mutable] [color_hash #:mutable]))
(struct colAttr ([width #:mutable] [back_color #:mutable]))

(struct line-chart-sheet ([topic #:mutable] [x_data_range #:mutable] [y_data_range_list #:mutable]))
(struct data-range ([sheet_name #:mutable] [range_str #:mutable]))
(struct data-serial ([topic #:mutable] [data_range #:mutable]))

(define xlsx%
  (class object%
         (super-new)
         
         (field 
          [sheets '()]
          )
         
         (define/public (add-data-sheet sheet_name sheet_data)
           (let* ([sheet_length (length sheets)]
                  [seq (add1 sheet_length)]
                  [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'data)) sheets)))])
             (set! sheets `(,@sheets
                            ,(sheet
                              sheet_name
                              seq
                              'data
                              type_seq
                              (data-sheet sheet_data (make-hash) (make-hash)))))))

         (define/public (add-line-chart-sheet sheet_name topic)
           (let* ([sheet_length (length sheets)]
                  [seq (add1 sheet_length)]
                  [type_seq (add1 (length (filter (lambda (rec) (eq? (sheet-type rec) 'chart)) sheets)))])
             (set! sheets `(,@sheets
                            ,(sheet
                              sheet_name
                              seq
                              'chart
                              type_seq
                              (line-chart-sheet topic (data-range "" "") '()))))))

         (define/public (set-line-chart-x-data! line_chart_sheet sheet_name data_range)
           (set-line-chart-sheet-x_data_range! line_chart_sheet (data-range sheet_name data_range)))

         (define/public (add-line-chart-y-data! line_chart_sheet y_topic sheet_name data_range)
           (set-line-chart-sheet-y_data_range_list! line_chart_sheet `(,@(line-chart-sheet-y_data_range_list line_chart_sheet) ,(data-serial y_topic (data-range sheet_name data_range)))))
         
         (define/public (sheet-ref sheet_seq)
           (list-ref sheets sheet_seq))
         
         (define/public (set-sheet-col-width! sheet col_range width)
           (hash-set! (data-sheet-width_hash (sheet-content sheet)) col_range width))

         (define/public (set-sheet-col-color! sheet col_range color)
           (hash-set! (data-sheet-color_hash (sheet-content sheet)) col_range color))
         ))

