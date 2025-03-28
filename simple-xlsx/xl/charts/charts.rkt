#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt"
         "./charts-lib.rkt"
         "./line-chart.rkt"
         "./pie-chart.rkt"
         "./bar-chart.rkt")

(provide (contract-out
          [to-chart-head (-> list?)]
          [to-chart-title (-> string? list?)]
          [from-chart-title (-> hash? void?)]
          [to-plotArea-head (-> list)]
          [to-chart-line (-> list?)]
          [to-chart-3d-line (-> list?)]
          [to-chart-pie (-> list?)]
          [to-chart-3d-pie (-> list?)]
          [to-chart-bar (-> list?)]
          [to-chart-3d-bar (-> list?)]
          [to-chart (-> list?)]
          [from-chart (-> hash? void?)]
          [write-charts (->* () (path-string?) void?)]
          [read-charts (->* () (path-string?) void?)]
          ))

(define (to-chart-head)
  '(
    "c:chartSpace"
    ("xmlns:c" . "http://schemas.openxmlformats.org/drawingml/2006/chart")
    ("xmlns:a" . "http://schemas.openxmlformats.org/drawingml/2006/main")
    ("xmlns:r" . "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ("c:lang" ("val" . "zh-CN"))))

(define (to-chart-title topic)
  (list
   "c:title"
   (list "c:tx"
         (list "c:rich"
               (list "a:bodyPr")
               (list "a:lstStyle")
               (list "a:p"
                     (list "a:pPr"
                           (list "a:defRPr"))
                     (list "a:r"
                           (list "a:rPr" (cons "lang" "zh-CN") (cons "altLang" "en-US"))
                           (list "a:t" topic))
                     (list "a:endParaRPr" (cons "lang" "en-US") (cons "altLang" "zh-CN")))))
   (list "c:layout")))

(define (from-chart-title xml_hash)
  (if (hash-has-key? xml_hash "c:chartSpace1.c:chart1.c:title1.c:tx1.c:rich1.a:p1.a:r1.a:t1")
      (set-CHART-SHEET-topic! (*CURRENT_SHEET*) (hash-ref xml_hash "c:chartSpace1.c:chart1.c:title1.c:tx1.c:rich1.a:p1.a:r1.a:t1"))
      (set-CHART-SHEET-topic! (*CURRENT_SHEET*) "")))

(define (to-plotArea-head)
  '(
    "c:plotArea"
    ("c:layout")))

(define (to-chart-line)
  `(
    ,@(to-chart-head)
    ,(list "c:chart"
           (to-chart-title (CHART-SHEET-topic (*CURRENT_SHEET*)))
           (append
            (to-plotArea-head)
            (list (to-line-chart-sers (CHART-SHEET-serial (*CURRENT_SHEET*))))
            (list (catAx))
            (list (valAx)))
           (legend)
           (plot-vis-only))))

(define (to-chart-3d-line)
  `(
    ,@(to-chart-head)
    ,(list "c:chart"
           (to-chart-title (CHART-SHEET-topic (*CURRENT_SHEET*)))
           (view3d)
           (append
            (to-plotArea-head)
            (list (to-line-3d-chart-sers (CHART-SHEET-serial (*CURRENT_SHEET*))))
            (list (catAx))
            (list (valAx)))
           (legend)
           (plot-vis-only))))

(define (to-chart-pie)
  `(
    ,@(to-chart-head)
    ,(list "c:chart"
           (to-chart-title (CHART-SHEET-topic (*CURRENT_SHEET*)))
           (append
            (to-plotArea-head)
            (list (to-pie-chart-sers (CHART-SHEET-serial (*CURRENT_SHEET*)))))
           (legend)
           (plot-vis-only))))

(define (to-chart-3d-pie)
  `(
    ,@(to-chart-head)
    ,(list "c:chart"
           (to-chart-title (CHART-SHEET-topic (*CURRENT_SHEET*)))
           (view3d)
           (append
            (to-plotArea-head)
            (list (to-pie-3d-chart-sers (CHART-SHEET-serial (*CURRENT_SHEET*)))))
           (legend)
           (plot-vis-only))))

(define (to-chart-bar)
  `(
    ,@(to-chart-head)
    ,(list "c:chart"
           (to-chart-title (CHART-SHEET-topic (*CURRENT_SHEET*)))
           (append
            (to-plotArea-head)
            (list (to-bar-chart-sers (CHART-SHEET-serial (*CURRENT_SHEET*))))
            (list (catAx))
            (list (valAx)))
           (legend)
           (plot-vis-only))))

(define (to-chart-3d-bar)
  `(
    ,@(to-chart-head)
    ,(list "c:chart"
           (to-chart-title (CHART-SHEET-topic (*CURRENT_SHEET*)))
           (view3d)
           (append
            (to-plotArea-head)
            (list (to-bar-3d-chart-sers (CHART-SHEET-serial (*CURRENT_SHEET*))))
            (list (catAx))
            (list (valAx)))
           (legend)
           (plot-vis-only))))

(define (to-chart)
  (cond
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE)
    (to-chart-line)]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'LINE3D)
    (to-chart-3d-line)]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE)
    (to-chart-pie)]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'PIE3D)
    (to-chart-3d-pie)]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'BAR)
    (to-chart-bar)]
   [(eq? (CHART-SHEET-chart_type (*CURRENT_SHEET*)) 'BAR3D)
    (to-chart-3d-bar)]
   ))

(define (from-chart xml_hash)
  (cond
   [(hash-has-key? xml_hash "c:chartSpace1.c:chart1.c:plotArea1.c:lineChart1")
    (set-CHART-SHEET-chart_type! (*CURRENT_SHEET*) 'LINE)]
   [(hash-has-key? xml_hash "c:chartSpace1.c:chart1.c:plotArea1.c:line3DChart1")
    (set-CHART-SHEET-chart_type! (*CURRENT_SHEET*) 'LINE3D)]
   [(hash-has-key? xml_hash "c:chartSpace1.c:chart1.c:plotArea1.c:barChart1")
    (set-CHART-SHEET-chart_type! (*CURRENT_SHEET*) 'BAR)]
   [(hash-has-key? xml_hash "c:chartSpace1.c:chart1.c:plotArea1.c:bar3DChart1")
    (set-CHART-SHEET-chart_type! (*CURRENT_SHEET*) 'BAR3D)]
   [(hash-has-key? xml_hash "c:chartSpace1.c:chart1.c:plotArea1.c:pieChart1")
    (set-CHART-SHEET-chart_type! (*CURRENT_SHEET*) 'PIE)]
   [(hash-has-key? xml_hash "c:chartSpace1.c:chart1.c:plotArea1.c:pie3DChart1")
    (set-CHART-SHEET-chart_type! (*CURRENT_SHEET*) 'PIE3D)])

  (from-chart-title xml_hash)

  (set-CHART-SHEET-serial! (*CURRENT_SHEET*) (from-sers xml_hash)))

(define (write-charts [output_dir #f])
  (let loop ([sheets (XLSX-sheet_list (*XLSX*))]
             [sheet_index 0]
             [chart_sheet_index 1])
    (when (not (null? sheets))
      (if (CHART-SHEET? (car sheets))
          (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "charts"))])
            (make-directory* dir)

            (with-sheet-ref
             sheet_index
             (lambda ()
               (with-output-to-file (build-path dir (format "chart~a.xml" chart_sheet_index))
                 #:exists 'replace
                 (lambda ()
                   (printf (lists-to-xml (to-chart)))))))

            (loop (cdr sheets) (add1 sheet_index) (add1 chart_sheet_index)))
          (loop (cdr sheets) (add1 sheet_index) chart_sheet_index)))))

(define (read-charts [input_dir #f])
  (let ([dir (if input_dir input_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "charts"))])
    (let loop ([sheets (XLSX-sheet_list (*XLSX*))]
               [sheet_index 0]
               [chart_sheet_index 1])
      (when (not (null? sheets))
        (if (CHART-SHEET? (car sheets))
            (begin
              (with-sheet-ref
               sheet_index
               (lambda ()
                 (from-chart (xml-file-to-hash
                              (build-path dir (format "chart~a.xml" chart_sheet_index))
                              '(
                                "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:tx.c:v"
                                "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:cat.c:strRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:lineChart.c:ser.c:val.c:numRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:line3DChart.c:ser.c:tx.c:v"
                                "c:chartSpace.c:chart.c:plotArea.c:line3DChart.c:ser.c:cat.c:strRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:line3DChart.c:ser.c:val.c:numRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:tx.c:v"
                                "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:cat.c:strRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:barChart.c:ser.c:val.c:numRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:bar3DChart.c:ser.c:tx.c:v"
                                "c:chartSpace.c:chart.c:plotArea.c:bar3DChart.c:ser.c:cat.c:strRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:bar3DChart.c:ser.c:val.c:numRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:tx.c:v"
                                "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:cat.c:strRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:pieChart.c:ser.c:val.c:numRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:pie3DChart.c:ser.c:tx.c:v"
                                "c:chartSpace.c:chart.c:plotArea.c:pie3DChart.c:ser.c:cat.c:strRef.c:f"
                                "c:chartSpace.c:chart.c:plotArea.c:pie3DChart.c:ser.c:val.c:numRef.c:f"
                                "c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.a:r.a:t"
                                "c:chartSpace.c:chart.c:plotArea.c:lineChart"
                                "c:chartSpace.c:chart.c:plotArea.c:line3DChart"
                                "c:chartSpace.c:chart.c:plotArea.c:barChart"
                                "c:chartSpace.c:chart.c:plotArea.c:bar3DChart"
                                "c:chartSpace.c:chart.c:plotArea.c:pieChart"
                                "c:chartSpace.c:chart.c:plotArea.c:pie3DChart"
                                )))))
              (loop (cdr sheets) (add1 sheet_index) (add1 chart_sheet_index)))
            (loop (cdr sheets) (add1 sheet_index) chart_sheet_index))))))
