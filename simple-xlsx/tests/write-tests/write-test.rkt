#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(define write-test
  (test-suite
   "test-normal-data-sheet"

   (test-case 
    "simple1"
   
    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (let ([xlsx (new xlsx%)])
            (send xlsx add-data-sheet 
                  #:sheet_name "DataSheet" 
                  #:sheet_data '(("month/brand" "201601" "201602" "201603" "201604")
                                 ("CAT" 100 300 200 6.9)
                                 ("Puma" 200 400 300 13.89)
                                 ("Brooks" 300 500 400 23.34)
                                 ))
            (send xlsx set-data-sheet-col-width! #:sheet_name "DataSheet" #:col_range "A-B" #:width 50)
            (send xlsx set-data-sheet-cell-color! #:sheet_name "DataSheet" #:cell_range "B2-C3" #:color "FF0000")
            (send xlsx set-data-sheet-cell-color! #:sheet_name "DataSheet" #:cell_range "C4-D4" #:color "0000FF")

            (send xlsx add-chart-sheet #:sheet_name "LineChart1" #:topic "Horizontal Data" #:x_topic "Kg")
            (send xlsx set-chart-x-data! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B1-D1")
            (send xlsx add-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")
            (send xlsx add-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B3-D3" #:y_topic "Puma")
            (send xlsx add-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "DataSheet" #:data_range "B4-D4" #:y_topic "Brooks")

            (send xlsx add-chart-sheet #:sheet_name "LineChart2" #:topic "Vertical Data" #:x_topic "Kg")
            (send xlsx set-chart-x-data! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "A2-A4" )
            (send xlsx add-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "B2-B4" #:y_topic "201601")
            (send xlsx add-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "C2-C4" #:y_topic "201602")
            (send xlsx add-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "DataSheet" #:data_range "D2-D4" #:y_topic "201603")

            (send xlsx add-chart-sheet #:sheet_name "LineChart3D" #:chart_type 'line3d #:topic "LineChart3D" #:x_topic "Kg")
            (send xlsx set-chart-x-data! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "A2-A4" )
            (send xlsx add-chart-serial! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "B2-B4" #:y_topic "201601")
            (send xlsx add-chart-serial! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "C2-C4" #:y_topic "201602")
            (send xlsx add-chart-serial! #:sheet_name "LineChart3D" #:data_sheet_name "DataSheet" #:data_range "D2-D4" #:y_topic "201603")

            (send xlsx add-chart-sheet #:sheet_name "BarChart" #:chart_type 'bar #:topic "BarChart" #:x_topic "Kg")
            (send xlsx set-chart-x-data! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
            (send xlsx add-chart-serial! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")
            (send xlsx add-chart-serial! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B3-D3" #:y_topic "Puma")
            (send xlsx add-chart-serial! #:sheet_name "BarChart" #:data_sheet_name "DataSheet" #:data_range "B4-D4" #:y_topic "Brooks")

            (send xlsx add-chart-sheet #:sheet_name "BarChart3D" #:chart_type 'bar3d #:topic "BarChart3D" #:x_topic "Kg")
            (send xlsx set-chart-x-data! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
            (send xlsx add-chart-serial! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")
            (send xlsx add-chart-serial! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B3-D3" #:y_topic "Puma")
            (send xlsx add-chart-serial! #:sheet_name "BarChart3D" #:data_sheet_name "DataSheet" #:data_range "B4-D4" #:y_topic "Brooks")

            (send xlsx add-chart-sheet #:sheet_name "PieChart" #:chart_type 'pie #:topic "PieChart" #:x_topic "Kg")
            (send xlsx set-chart-x-data! #:sheet_name "PieChart" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
            (send xlsx add-chart-serial! #:sheet_name "PieChart" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")

            (send xlsx add-chart-sheet #:sheet_name "PieChart3D" #:chart_type 'pie3d #:topic "PieChart3D" #:x_topic "Kg")
            (send xlsx set-chart-x-data! #:sheet_name "PieChart3D" #:data_sheet_name "DataSheet" #:data_range "B1-D1" )
            (send xlsx add-chart-serial! #:sheet_name "PieChart3D" #:data_sheet_name "DataSheet" #:data_range "B2-D2" #:y_topic "CAT")

            (write-xlsx-file xlsx "test.xlsx")
            )

          (with-input-from-xlsx-file
           "test.xlsx"
           (lambda (xlsx)
             (check-equal? (get-sheet-names xlsx) '("DataSheet" "LineChart1" "LineChart2" "LineChart3D" "BarChart" "BarChart3D" "PieChart" "PieChart3D"))

             (load-sheet "DataSheet" xlsx)
             (check-equal? (get-sheet-dimension xlsx) '(4 . 5))
             (check-equal? (get-cell-value "E2" xlsx) 6.9)

             (load-sheet "LineChart1" xlsx)
             (check-equal? (get-sheet-dimension xlsx) '(4 . 5))

             ))
          )
        (lambda () 
          (delete-file "test.xlsx"))
        ))
   ))

(run-tests write-test)
