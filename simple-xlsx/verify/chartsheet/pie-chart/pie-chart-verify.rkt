#lang racket

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path pie_chart_file "_pie_chart.xlsx")
(define-runtime-path pie_chart_read_and_write_file "_pie_chart_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-piechart"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           pie_chart_file
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
              "PieChart" 'PIE "PieChartExample"
              '(
                ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
                ))))

          (read-and-write-xlsx
           pie_chart_file
           pie_chart_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
          ;(void)
          (delete-file pie_chart_file)
          (delete-file pie_chart_read_and_write_file)
          )))
   ))

(run-tests test-writer)
