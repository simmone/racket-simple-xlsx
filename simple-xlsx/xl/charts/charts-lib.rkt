#lang racket

(require "../../lib/dimension.rkt"
         "../../lib/sheet-lib.rkt"
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt")

(provide (contract-out
          [to-ser-head (-> natural? string? list?)]
          [from-ser-head (-> hash? natural? string?)]
          [to-ser-cat (-> string? string? list?)]
          [from-ser-cat (-> hash? natural? (cons/c string? string?))]
          [to-ser-val (-> string? string? list?)]
          [from-ser-val (-> hash? natural? (cons/c string? string?))]
          [to-ser (-> natural? (list/c string? string? string? string? string?) list?)]
          [from-ser (-> hash? natural? (list/c string? string? string? string? string?))]
          [to-sers (-> (listof (list/c string? string? string? string? string?)) list?)]
          [from-sers (-> hash? (listof (list/c string? string? string? string? string?)))]
          [catAx (-> list?)]
          [valAx (-> list?)]
          [legend (-> list?)]
          [plot-vis-only (-> list?)]
          [marker-axid-tail (-> list?)]
          [view3d (-> list?)]
          ))

(define (to-ser-head ser_index x_title)
  (list
   "c:ser"
   (list "c:idx" (cons "val" (number->string ser_index)))
   (list "c:order" (cons "val" (number->string ser_index)))
   (list "c:tx"
         (list "c:v" x_title))
   '("c:marker"
     ("c:symbol" ("val" . "none")))))

(define (get-chart-type-node)
  (cond
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE)
    "c:lineChart"]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE3D)
    "c:line3DChart"]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE)
    "c:pieChart"]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE3D)
    "c:pie3DChart"]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'BAR)
    "c:barChart"]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'BAR3D)
    "c:bar3DChart"]))

(define (from-ser-head xml_hash ser_index)
  (let* ([key_head "c:chartSpace1.c:chart1.c:plotArea1"]
         [chart_type (get-chart-type-node)]
         [key_suffix (format "c:ser~a.c:tx1.c:v1" ser_index)]
         [key (format "~a.~a1.~a" key_head chart_type key_suffix)])
    (hash-ref xml_hash key "")))

(define (to-ser-cat sheet_name range)
  (ser-sec 'cat sheet_name range))

(define (from-ser-cat xml_hash ser_index)
  (let* ([key_head "c:chartSpace1.c:chart1.c:plotArea1"]
         [chart_type (get-chart-type-node)]
         [key_suffix (format "c:ser~a.c:cat1.c:strRef1.c:f1" ser_index)]
         [key (format "~a.~a1.~a" key_head chart_type key_suffix)]
         [sheet_ref (hash-ref xml_hash key #f)])

    (if sheet_ref
        (let ([ref_pair (regexp-split #rx"!" sheet_ref)])
          (cons
           (first ref_pair)
           (range_xml->range (second ref_pair))))
        '("" . ""))))

(define (to-ser-val sheet_name range)
  (ser-sec 'val sheet_name range))

(define (from-ser-val xml_hash ser_index)
  (let* ([key_head "c:chartSpace1.c:chart1.c:plotArea1"]
         [chart_type (get-chart-type-node)]
         [key_suffix (format "c:ser~a.c:val1.c:numRef1.c:f1" ser_index)]
         [key (format "~a.~a1.~a" key_head chart_type key_suffix)]
         [sheet_ref (hash-ref xml_hash key #f)])

    (if sheet_ref
        (let ([ref_pair (regexp-split #rx"!" sheet_ref)])
          (cons
           (first ref_pair)
           (range_xml->range (second ref_pair))))
        '("" . ""))))

(define (from-ser xml_hash ser_index)
  (let ([ser_head (from-ser-head xml_hash ser_index)]
        [ser_category (from-ser-cat xml_hash ser_index)]
        [ser_value (from-ser-val xml_hash ser_index)])

    (list
     ser_head
     (car ser_category)
     (cdr ser_category)
     (car ser_value)
     (cdr ser_value))))

(define (get-sheet-name-range-values sheet_name range)
  (with-sheet-name
   sheet_name
   (lambda ()
     (get-range-values range))))

(define (ser-sec type sheet_name range)
  (list
   (format "c:~a" type)
   (list
    (format "c:~aRef" (if (eq? type 'cat) 'str 'num))
    (list
     "c:f"
     (format "~a!~a" sheet_name (range->range_xml range)))
    (append
     (list (format "c:~aCache" (if (eq? type 'cat) 'str 'num)))
     (let ([cat_cell_values (get-sheet-name-range-values sheet_name range)])
       (append
        (list (list "c:ptCount" (cons "val" (number->string (length cat_cell_values)))))
        (let loop ([cell_values cat_cell_values]
                   [index 0]
                   [result_list '()])
          (if (not (null? cell_values))
              (loop
               (cdr cell_values)
               (add1 index)
               (cons
                (list
                 "c:pt"
                 (cons "idx" (number->string index))
                 (list
                  "c:v" (format "~a" (car cell_values))))
                result_list))
              (reverse result_list)))))))))

(define (to-ser index serial)
  (append
   (to-ser-head index (first serial))
   (list (to-ser-cat (second serial) (third serial)))
   (list (to-ser-val (fourth serial) (fifth serial)))))

(define (to-sers ser_list)
  (let loop ([_sers ser_list]
             [ser_index 0]
             [result_list '()])
    (if (not (null? _sers))
        (loop
         (cdr _sers)
         (add1 ser_index)
         (cons (to-ser ser_index (car _sers)) result_list))
        (reverse result_list))))

(define (from-sers xml_hash)
  (let* ([key_head "c:chartSpace1.c:chart1.c:plotArea1"]
         [chart_type (get-chart-type-node)]
         [key_suffix "c:ser's count"]
         [key (format "~a.~a1.~a" key_head chart_type key_suffix)]
         [sers_count (hash-ref xml_hash key 0)])
    
    (let loop ([ser_index 1]
               [sers_list '()])
      (if (<= ser_index sers_count)
          (loop
           (add1 ser_index)
           (cons
            (from-ser xml_hash ser_index)
            sers_list))
          (reverse sers_list)))))

(define (catAx)
  '("c:catAx"
    ("c:axId" ("val" . "76367360"))
    ("c:scaling"
     ("c:orientation" ("val" . "minMax")))
    ("c:axPos" ("val" . "b"))
    ("c:numFmt" ("formatCode" . "General") ("sourceLinked" . "1"))
    ("c:majorTickMark" ("val" . "none"))
    ("c:tickLblPos" ("val" . "nextTo"))
    ("c:crossAx" ("val" . "76368896"))
    ("c:crosses" ("val" . "autoZero"))
    ("c:auto" ("val" . "1"))
    ("c:lblAlgn" ("val" . "ctr"))
    ("c:lblOffset" ("val" . "100"))))

(define (valAx)
  '("c:valAx"
    ("c:axId" ("val". "76368896"))
    ("c:scaling"
     ("c:orientation" ("val" . "minMax")))
    ("c:axPos" ("val" . "l"))
    ("c:majorGridlines")
    ("c:title"
     ("c:tx"
      ("c:rich"
       ("a:bodyPr")
       ("a:lstStyle")
       ("a:p"
        ("a:pPr"
         ("a:defRPr"))
        ("a:r"
         ("a:rPr" ("lang" . "zh-CN") ("altLang" . "en-US"))
         ("a:t"))
        ("a:endParaRPr" ("lang". "en-US") ("altLang" . "zh-CN")))))
     ("c:layout"))
    ("c:numFmt" ("formatCode" . "General") ("sourceLinked" . "1"))
    ("c:majorTickMark" ("val" . "none"))
    ("c:tickLblPos" ("val". "nextTo"))
    ("c:crossAx" ("val" . "76367360"))
    ("c:crosses" ("val" . "autoZero"))
    ("c:crossBetween" ("val" . "between"))))

(define (legend)
  '("c:legend"
    ("c:legendPos" ("val" . "r"))
    ("c:layout")))

(define (plot-vis-only)
  '("c:plotVisOnly" ("val" . "1")))

(define (marker-axid-tail)
  '(
    ("c:marker" ("val" . "1"))
    ("c:axId" ("val" . "76367360"))
    ("c:axId" ("val" . "76368896"))))

(define (view3d)
  '("c:view3D" ("c:perspective" ("val" . "30"))))
