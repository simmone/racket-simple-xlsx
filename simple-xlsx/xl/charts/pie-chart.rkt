#lang racket

(require "lib.rkt")
(require "../../sheet/sheet.rkt")

(provide (contract-out
          [pie-chart-head (-> list?)]
          [pie-3d-chart-head (-> list?)]
          [pie-chart-tail (-> list?)]
          [to-pie-chart-sers (-> (listof (list/c string? string? string? string? string?)) list?)]
          [to-pie-3d-chart-sers (-> (listof (list/c string? string? string? string? string?)) list?)]
          ))

(define (pie-chart-head)
  '("c:pieChart"
    ("c:varyColors" ("val" . "1"))))

(define (pie-3d-chart-head)
  '("c:pie3DChart"
    ("c:varyColors" ("val" . "1"))))

(define (pie-chart-tail)
  '(
    ("c:firstSliceAng" ("val" . "0"))))

(define (to-pie-chart-sers ser_list)
  (append
   (pie-chart-head)
   (list (car (to-sers ser_list)))
   (pie-chart-tail)))

(define (to-pie-3d-chart-sers ser_list)
  (append
   (pie-3d-chart-head)
   (list (car (to-sers ser_list)))
   (pie-chart-tail)))
