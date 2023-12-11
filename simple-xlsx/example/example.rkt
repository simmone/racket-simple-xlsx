#lang racket

(require "../main.rkt")

(require racket/date)

(let ([sheet_data (list
                   (list "month/brand" "201601" "201602" "201603" "201604" "201605")
                   (list "CAT" 100 300 200 0.6934 (seconds->date (find-seconds 0 0 0 17 9 2018)))
                   (list "Puma" 200 400 300 139999.89223 (seconds->date (find-seconds 0 0 0 18 9 2018)))
                   (list "Asics" 300 500 400 23.34 (seconds->date (find-seconds 0 0 0 19 9 2018))))]
      [sheet_data2 (list
                    (list "month/brand" "201601" "201602" "201603" "201604" "201605" "alignment")
                    (list "CAT" 100 300 200 0.6934 (seconds->date (find-seconds 0 0 0 17 9 2018)) "")
                    (list "Puma" 200 400 300 -139999.89223 (seconds->date (find-seconds 0 0 0 18 9 2018)) "")
                    (list "Asics" 300 500 400 23.34 (seconds->date (find-seconds 0 0 0 19 9 2018)) "")
                    (list "" "" "" "" "" "" "Left")
                    (list "" "" "" "" "" "" "Right")
                    (list "" "" "" "" "" "" "Center")
                    (list "" "" "" "" "" "" "Top")
                    (list "" "" "" "" "" "" "Bottom")
                    (list "" "" "" "" "" "" "Middle")
                    (list "" "" "" "" "" "" "Center/Middle")
                    )])

  (write-xlsx
   "test.xlsx"
   (lambda ()
     (add-data-sheet "DataSheet" sheet_data)
     (with-sheet-ref
      0
      (lambda ()
        (set-col-range-width "A-B" 50)
        (set-freeze-row-col-range 1 1)))

     (add-data-sheet "DataSheetWithStyle" sheet_data)
     (with-sheet-ref
      1
      (lambda ()
        (set-col-range-width "A-B" 50)
        (set-row-range-height "3-4" 30)
        (set-col-range-width "F" 20)
        (set-cell-range-fill-style "A2-B3" "00C851" "solid")
        (set-cell-range-fill-style "C3-D4" "AA66CC" "solid")
        (set-cell-range-font-style "B3-C4" 20 "Impact" "000000")
        (set-cell-range-font-style "B1-C3" 20 "Impact" "FF8800")
        (set-cell-range-number-style "E2-E2" "0.00%")
        (set-cell-range-number-style "E3-E3" "0,000.00")
        (set-cell-range-number-style "E4-E4" "0")
        (set-cell-range-border-style "B2-C4" "all" "0000FF" "dashed")
        (set-cell-range-date-style "F2-F2" "yyyy-mm-dd")
        (set-cell-range-date-style "F3-F3" "yyyy/mm/dd")
        (set-cell-range-date-style "F4-F4" "yyyy年mm月dd日")))

     (add-data-sheet "DataSheetWithStyle2" sheet_data2)
     (with-sheet-ref
      2
      (lambda ()
        (set-col-range-width "1-1" 20)
        (set-row-range-height "2-4" 30)
        (set-col-range-width "2-6" 10)
        (set-row-range-fill-style "1-3" "00C851" "solid")
        (set-col-range-fill-style "1-6" "AA66CC" "solid")
        (set-cell-range-fill-style "B1-C3" "FF8800" "solid")
        (set-col-range-width "7" 50)
        (set-row-range-height "5-11" 50)
        (set-cell-range-alignment-style "G5" "left" "center")
        (set-cell-range-alignment-style "G6" "right" "center")
        (set-cell-range-alignment-style "G7" "center" "center")
        (set-cell-range-alignment-style "G8" "center" "top")
        (set-cell-range-alignment-style "G9" "center" "bottom")
        (set-cell-range-alignment-style "G10" "center" "center")
        (set-cell-range-alignment-style "G11" "center" "center")

        (set-col-range-width "E-F" 30)
        (set-cell-range-number-style "G1" "@__&quot;means&quot;__@")
        (set-col-range-number-style "E-F" "￥#,##0.00;[Red]￥-#,##0.00")))

     (add-chart-sheet
      "LineChart" 'LINE "LineChart"
      '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Books" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))

     (add-chart-sheet
      "Line3DChart" 'LINE3D "Line3DChartExample"
      '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Books" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))

     (add-chart-sheet
      "BarChart" 'BAR "BarChart"
      '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Books" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))

     (add-chart-sheet
      "Bar3DChart" 'BAR3D "Bar3DChart"
      '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")
        ("Puma" "DataSheet" "B1-D1" "DataSheet" "B3-D3")
        ("Books" "DataSheet" "B1-D1" "DataSheet" "B4-D4")))

     (add-chart-sheet
      "PieChart" 'PIE "PieChart"
      '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")))

     (add-chart-sheet
      "Pie3DCart" 'PIE3D "Pie3DChart"
      '(
        ("CAT" "DataSheet" "B1-D1" "DataSheet" "B2-D2")))))

  (read-xlsx
   "test.xlsx"
   (lambda ()
     (printf "~a\n" (get-sheet-name-list))
     ;("DataSheet" "LineChart1" "LineChart2" "LineChart3D" "BarChart" "BarChart3D" "PieChart" "PieChart3D"))

     (with-sheet
      (lambda ()
        (printf "~a\n" (get-sheet-dimension)) ; A1:F4

        (printf "~a\n" (get-cell "A2")) ;CAT

        (let ([date_val (oa_date_number->date (get-cell "F2"))])
          (printf "~a,~a,~a\n" (date-year date_val) (date-month date_val) (date-day date_val))) ; 2018,9,17

        (printf "~a\n" (get-rows)
        ; ((month/brand 201601 201602 201603 201604 201605) (CAT 100 300 200 0.6934 43360) (Puma 200 400 300 139999.89223 43361) (Asics 300 500 400 23.34 43362))
                )))))

  (read-and-write-xlsx
   "test.xlsx"
   "write_back.xlsx"
   (lambda ()
     (void))))
