#lang racket

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         racket/runtime-path)

(define-runtime-path line_chart_file "_line_chart.xlsx")
(define-runtime-path line_chart_read_and_write_file "_line_chart_read_and_write.xlsx")

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
                ("" "201601" "201602" "201603" "201604")
                ("CAT" 100 300 200 400)
                ("DOG" 200 400 300 100)
                ("RABBIT" 300 500 400 200)
                ))

             (add-chart-sheet
              "LineChart" 'LINE "LineChartExample"
              '(
                ("CAT" "DataSheet" "B1-E1" "DataSheet" "B2-E2")
                ("Puma" "DataSheet" "B1-E1" "DataSheet" "B3-E3")
                ("Brooks" "DataSheet" "B1-E1" "DataSheet" "B4-E4")
               ))
             ))

          (read-and-write-xlsx
           line_chart_file
           line_chart_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
          ;(void)
          (delete-file line_chart_file)
          (delete-file line_chart_read_and_write_file)
          )))
   ))

(run-tests test-writer)
