#lang racket

(require "../../../xlsx.rkt")

(provide (contract-out
          [write-line-chart (-> (is-a?/c xlsx%) string? list? string?)]
          ))

(define (print-x-data xlsx x_data_range)
  (with-output-to-string
    (lambda ()
      (let* ([sheet_name (data-range-sheet_name x_data_range)]
             [range_str (data-range-range_str x_data_range)])

        (printf "<c:cat><c:strRef><c:f>~a!~a</c:f><c:strCache><c:ptCount val=\"~a\"/>"
                sheet_name
                (convert-range range_str)
                (range-length range_str))

        (let loop ([loop_list (send xlsx get-range-data sheet_name range_str)]
                   [idx 0])
          (when (not (null? loop_list))
                (printf "<c:pt idx=\"~a\"><c:v>~a</c:v></c:pt>" idx (car loop_list))
                (loop (cdr loop_list) (add1 idx))))
        (printf "</c:strCache></c:strRef></c:cat>")))))

(define (write-line-chart xlsx x_data_range y_data_range_list)
  (with-output-to-string
    (lambda ()
      (printf "<c:lineChart><c:grouping val=\"standard\"/>")
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
                (printf "<c:marker><c:symbol val=\"none\"/></c:marker>")
                (printf "~a<c:val><c:numRef>" (print-x-data xlsx x_data_range))
                (printf "<c:f>~a!~a</c:f>" sheet_name (convert-range range_str))
                (printf "<c:numCache>")
                (printf "<c:ptCount val=\"~a\"/>" (range-length range_str))
                (let y-data-loop ([y_data_loop_list (send xlsx get-range-data sheet_name range_str)]
                                  [idx 0])
                  (when (not (null? y_data_loop_list))
                        (printf "<c:pt idx=\"~a\"><c:v>~a</c:v></c:pt>" idx (car y_data_loop_list))
                        (y-data-loop (cdr y_data_loop_list) (add1 idx))))
                (printf "</c:numCache></c:numRef></c:val></c:ser>"))
              (loop (cdr loop_list) (add1 ser_seq))))
      (printf "<c:marker val=\"1\"/><c:axId val=\"76367360\"/><c:axId val=\"76368896\"/></c:lineChart>")
      )))

