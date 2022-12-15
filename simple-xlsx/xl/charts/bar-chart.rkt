#lang racket

(require "lib.rkt")
(require "../../sheet/sheet.rkt")

(provide (contract-out
          [bar-chart-head (-> list?)]
          [bar-3d-chart-head (-> list?)]
          [to-bar-chart-sers (-> (listof (list/c string? string? string? string? string?)) list?)]
          [to-bar-3d-chart-sers (-> (listof (list/c string? string? string? string? string?)) list?)]
          ))

(define (bar-chart-head)
  '("c:barChart"
    ("c:barDir" ("val" . "col"))
    ("c:grouping" ("val" . "clustered"))))

(define (bar-3d-chart-head)
  '("c:bar3DChart"
    ("c:barDir" ("val" . "col"))
    ("c:grouping" ("val" . "clustered"))))

(define (to-bar-chart-sers ser_list)
  (append
   (bar-chart-head)
   (to-sers ser_list)
   (marker-axid-tail)))

(define (to-bar-3d-chart-sers ser_list)
  (append
   (bar-3d-chart-head)
   (to-sers ser_list)
   (marker-axid-tail)))
