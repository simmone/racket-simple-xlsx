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
                   (x_data list?)
                   (y_data_list list?)
                   )]
          [struct data-serial
                  (
                   (topic string?)
                   (data_list list?)
                   )]
          ))

(struct sheet ([name #:mutable] [seq #:mutable] [type #:mutable] [typeSeq #:mutable] [content #:mutable]))

(struct data-sheet ([rows #:mutable] [width_hash #:mutable] [color_hash #:mutable]))
(struct colAttr ([width #:mutable] [back_color #:mutable]))

(struct line-chart-sheet ([topic #:mutable] [x_data #:mutable] [y_data_list #:mutable]))
(struct data-serial ([topic #:mutable] [data_list #:mutable]))

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
                              (line-chart-sheet topic '() '()))))))

         (define/public (set-line-chart-x-data! line_chart_sheet data_list)
           (set-line-chart-sheet-x_data! line_chart_sheet data_list))

         (define/public (add-line-chart-y-data! line_chart_sheet y_topic data_list)
           (set-line-chart-sheet-y_data_list! line_chart_sheet `(,@(line-chart-sheet-y_data_list line_chart_sheet) ,(data-serial y_topic data_list))))
         
         (define/public (sheet-ref sheet_seq)
           (list-ref sheets sheet_seq))
         
         (define/public (set-sheet-col-width! sheet col_range width)
           (hash-set! (data-sheet-width_hash (sheet-content sheet)) col_range width))

         (define/public (set-sheet-col-color! sheet col_range color)
           (hash-set! (data-sheet-color_hash (sheet-content sheet)) col_range color))
         ))

