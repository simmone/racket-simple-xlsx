#lang racket

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path bar_chart_file "_bar_chart.xlsx")
(define-runtime-path bar_chart_read_and_write_file "_bar_chart_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-barchart"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           bar_chart_file
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
              "BarChart" 'BAR "BarChartExample"
              '(
                ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
                ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")
                ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
                ))))

          (read-and-write-xlsx
           bar_chart_file
           bar_chart_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
          ;(void)
          (delete-file bar_chart_file)
          (delete-file bar_chart_read_and_write_file)
          )))
   ))

(run-tests test-writer)
