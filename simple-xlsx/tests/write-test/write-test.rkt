#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(define write-test
  (test-suite
   "test-normal-data-sheet"

   (test-case 
    "simple1"
              
    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet 
            #:sheet_name "Sheet1" 
            #:sheet_data '(("201601" "201602" "201603" "201604" "201605" "201606" "201607" "201608" "201609" "201610" "201611" "201612")
                           (100 200 300 400 300 200 100 200 300 400 300 200)
                           (200 300 400 300 200 100 200 300 400 300 200 100)
                           (100 200 300 400 300 200 100 200 300 400 300 200)
                           (200 300 400 300 200 100 200 300 400 300 200 100)
                           (100 200 300 400 300 200 100 200 300 400 300 200)
                           (200 300 400 300 200 100 200 300 400 300 200 100)
                           (100 200 300 400 300 200 100 200 300 400 300 200)
                           (200 300 400 300 200 100 200 300 400 300 200 100)
                           (100 200 300 400 300 200 100 200 300 400 300 200)
                           (200 300 400 300 200 100 200 300 400 300 200 100)
                           (100 200 300 400 300 200 100 200 300 400 300 200)
                           (200 300 400 300 200 100 200 300 400 300 200 100)
                           ))

      (send xlsx add-line-chart-sheet #:sheet_name "LineChart1" #:topic "Horizontal Data" #:x_topic "Kg")
      (send xlsx set-line-chart-x-data! #:sheet_name "LineChart1" #:data_sheet_name "Sheet1" #:data_range "A1-L1")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "Sheet1" #:data_range "A2-L2" #:y_topic "2")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart1" #:data_sheet_name "Sheet1" #:data_range "A3-L3" #:y_topic "3")

      (send xlsx add-line-chart-sheet #:sheet_name "LineChart2" #:topic "Vertical Data" #:x_topic "Kg")
      (send xlsx set-line-chart-x-data! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "A1-L1" )
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "A2-A13" #:y_topic "A")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "B2-B13" #:y_topic "B")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "C2-C13" #:y_topic "C")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "D2-D13" #:y_topic "D")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "E2-E13" #:y_topic "E")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "F2-F13" #:y_topic "F")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "G2-G13" #:y_topic "G")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "H2-H13" #:y_topic "H")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "I2-I13" #:y_topic "I")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "J2-J13" #:y_topic "J")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "K2-K13" #:y_topic "K")
      (send xlsx add-line-chart-serial! #:sheet_name "LineChart2" #:data_sheet_name "Sheet1" #:data_range "L2-L13" #:y_topic "L")

      (write-xlsx-file xlsx "test1.xlsx")
      ))
   ))

(run-tests write-test)
