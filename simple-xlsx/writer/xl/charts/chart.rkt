#lang racket

(require "../../../xlsx.rkt")

(provide (contract-out
          [write-chart-sheet (-> string? (is-a?/c xlsx%) string?)]
          [write-chart-sheet-file (-> path-string? string? (is-a?/c xlsx%) exact-nonnegative-integer? void?)]
          ))

(define (print-x-data xlsx x_data_range)
  (with-output-to-string
    (lambda ()
      (let* ([sheet_name (data-range-sheet_name x_data_range)]
             [range_str (data-range-range_str x_data_range)])

        (printf "<c:cat><c:numRef><c:f>~a!~a</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"~a\"/>"
                sheet_name
                (convert-range range_str)
                (range-length range_str))

        (let loop ([loop_list (send xlsx get-range-data sheet_name range_str)]
                   [idx 0])
          (when (not (null? loop_list))
                (printf "<c:pt idx=\"~a\"><c:v>~a</c:v></c:pt>" idx (car loop_list))
                (loop (cdr loop_list) (add1 idx))))
        (printf "</c:numCache></c:numRef></c:cat>")))))

(define (print-y-data-list xlsx y_data_range_list)
  (with-output-to-string
    (lambda ()
      (let loop ([loop_list y_data_range_list]
                 [ser_seq 0])
        (when (not (null? loop_list))
              (let* ([y_data_serial (car loop_list)]
                     [topic (data-serial-topic y_data_serial)]
                     [data_range (data-serial-data_range y_data_serial)]
                     [sheet_name (data-range-sheet_name data_range)]
                     [range_str (data-range-range_str data_range)])
                (printf "<c:ser><c:idx val=\"~a\"/><c:order val=\"~a\"/>" ser_seq ser_seq)
                (printf "<c:tx><c:v>~a</c:v></c:tx>" topic)
                (printf "<c:marker><c:symbol val=\"none\"/></c:marker><c:val><c:numRef>")
                (printf "<c:f>~a!~a</c:f>" sheet_name (convert-range range_str))
                (printf "<c:numCache><c:formatCode>General</c:formatCode>")
                (printf "<c:ptCount val=\"~a\"/>" (range-length range_str))
                (let y-data-loop ([y_data_loop_list (send xlsx get-range-data sheet_name range_str)]
                                  [idx 0])
                  (when (not (null? y_data_loop_list))
                        (printf "<c:pt idx=\"~a\"><c:v>~a</c:v></c:pt>" idx (car y_data_loop_list))
                        (y-data-loop (cdr y_data_loop_list) (add1 idx))))
                (printf "</c:numCache></c:numRef></c:val></c:ser>"))
              (loop (cdr loop_list) (add1 ser_seq)))))))

(define (write-chart-sheet chart_sheet_name xlsx)
  (with-output-to-string
    (lambda ()
      (let* ([chart_sheet (sheet-content (send xlsx get-sheet-by-name chart_sheet_name))]
             [topic (line-chart-sheet-topic chart_sheet)]
             [x_data_range (line-chart-sheet-x_data_range chart_sheet)]
             [y_data_range_list (line-chart-sheet-y_data_range_list chart_sheet)])
        (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        (printf "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><c:lang val=\"zh-CN\"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang=\"zh-CN\" altLang=\"en-US\"/><a:t>~a</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title><c:plotArea><c:layout/><c:lineChart><c:grouping val=\"standard\"/>" topic)
        (printf "~a" (print-x-data xlsx x_data_range))
        (printf "~a" (print-y-data-list xlsx y_data_range_list))
        (printf "<c:marker val=\"1\"/><c:axId val=\"76367360\"/><c:axId val=\"76368896\"/></c:lineChart><c:catAx><c:axId val=\"76367360\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"b\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"76368896\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx><c:valAx><c:axId val=\"76368896\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"l\"/><c:majorGridlines/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"76367360\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx></c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>")))))

(define (write-chart-sheet-file dir sheet_name xlsx typeSeq)
  (with-output-to-file (build-path dir (format "chart~a.xml" typeSeq))
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-chart-sheet sheet_name xlsx)))))

