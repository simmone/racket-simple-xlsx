#lang racket

(require "lib.rkt")
(require "../../sheet/sheet.rkt")

(provide (contract-out
          [line-chart-head (-> list?)]
          [line-3d-chart-head (-> list?)]
          [to-line-chart-sers (-> (listof (list/c string? string? string? string? string?)) list?)]
          [to-line-3d-chart-sers (-> (listof (list/c string? string? string? string? string?)) list?)]
          ))

(define (line-chart-head)
  '("c:lineChart"
    ("c:grouping" ("val" . "standard"))))

(define (line-3d-chart-head)
  '("c:line3DChart"
    ("c:grouping" ("val" . "standard"))))

(define (to-line-chart-sers ser_list)
  (append
   (line-chart-head)
   (to-sers ser_list)
   (marker-axid-tail)))

(define (to-line-3d-chart-sers ser_list)
  (append
   (line-3d-chart-head)
   (to-sers ser_list)
   (marker-axid-tail)))


