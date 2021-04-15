#lang racket

(require "../../../xlsx/xlsx.rkt")
(require "../../../xlsx/sheet.rkt")

(require "line-chart.rkt")
(require "bar-chart.rkt")
(require "pie-chart.rkt")

(provide (contract-out
          [write-chart (-> string? (is-a?/c xlsx%) string?)]
          [write-chart-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define (write-chart chart_sheet_name xlsx)
  (with-output-to-string
    (lambda ()
      (let* ([chart_sheet (sheet-content (send xlsx get-sheet-by-name chart_sheet_name))]
             [chart_type (chart-sheet-chart_type chart_sheet)]
             [topic (chart-sheet-topic chart_sheet)]
             [x_topic (chart-sheet-x_topic chart_sheet)])
        (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        (printf "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><c:lang val=\"zh-CN\"/><c:chart>")
        (printf "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang=\"zh-CN\" altLang=\"en-US\"/><a:t>~a</a:t></a:r><a:endParaRPr lang=\"en-US\" altLang=\"zh-CN\"/></a:p></c:rich></c:tx><c:layout/></c:title>" topic)
        (cond
         [(or 
           (eq? chart_type 'line3d)
           (eq? chart_type 'bar3d))
          (printf "<c:view3D><c:perspective val=\"30\"/></c:view3D>")]
         [(eq? chart_type 'pie3d)
          (printf "<c:view3D><c:rotX val=\"40\"/></c:view3D>")]
         [else
          ""])
        (printf "<c:plotArea><c:layout/>")
        (cond
         [(or (eq? chart_type 'line) (eq? chart_type 'line3d))
          (printf "~a" (print-line-chart chart_sheet_name xlsx))]
         [(or (eq? chart_type 'bar) (eq? chart_type 'bar3d))
          (printf "~a" (print-bar-chart chart_sheet_name xlsx))]
         [(or (eq? chart_type 'pie) (eq? chart_type 'pie3d))
          (printf "~a" (print-pie-chart chart_sheet_name xlsx))]
         [else
          ""])
        (cond
         [(or 
           (eq? chart_type 'line) 
           (eq? chart_type 'line3d) 
           (eq? chart_type 'bar)
           (eq? chart_type 'bar3d)
           )
          (printf "<c:catAx><c:axId val=\"76367360\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"b\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"76368896\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx><c:valAx><c:axId val=\"76368896\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"l\"/><c:majorGridlines/>")
          (printf "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang=\"zh-CN\" altLang=\"en-US\"/><a:t>~a</a:t></a:r><a:endParaRPr lang=\"en-US\" altLang=\"zh-CN\"/></a:p></c:rich></c:tx><c:layout/></c:title>" x_topic)
          (printf "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"76367360\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx></c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>")]
         [else
          (printf "</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>")]
         )))))

(define (write-chart-file dir xlsx)
  (when (ormap (lambda (rec) (eq? (sheet-type rec) 'chart)) (get-field sheets xlsx))
        (make-directory* dir)

        (let loop ([loop_list (get-field sheets xlsx)])
          (when (not (null? loop_list))
                (when (eq? (sheet-type (car loop_list)) 'chart)
                      (with-output-to-file (build-path dir (format "chart~a.xml" (sheet-typeSeq (car loop_list))))
                        #:exists 'replace
                        (lambda ()
                          (printf "~a" (write-chart (sheet-name (car loop_list)) xlsx)))))
                (loop (cdr loop_list))))))

