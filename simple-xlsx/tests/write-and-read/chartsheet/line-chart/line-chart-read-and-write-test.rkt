#lang racket

(require rackunit/text-ui rackunit)

(require"../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path line_chart_file "line_chart.xlsx")
(define-runtime-path line_chart_read_and_write_file "line_chart_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-linechart"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           line_chart_file
           (lambda ()
             (add-data-sheet
              "DataSheet"
              '(
                ("201601" "201602" "201603" "201604")
                (100 300 200 400)
                (200 400 300 100)
                (300 500 400 200)
                ))

             (add-chart-sheet
              "LineChart" 'LINE "LineChartExample"
              '(
                ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
                ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")
                ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
               ))))

          (read-and-write-xlsx
           line_chart_file
           line_chart_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file line_chart_file)
          (delete-file line_chart_read_and_write_file)
          )))
   ))

(run-tests test-writer)
