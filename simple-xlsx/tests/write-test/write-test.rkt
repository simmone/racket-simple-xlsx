#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(define write-test
  (test-suite
   "test-normal-data-sheet"

   (test-case 
    "simple1"
              
    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "Sheet1" '(("201601" "201602" "201603" "201604" "201605" "201606" "201607" "201608" "201609" "201610" "201611" "201612")
                                           (100 200 300 400 300 200 100 200 300 400 300 200)
                                           (200 300 400 300 200 100 200 300 400 300 200 100)))

      (send xlsx add-line-chart-sheet "Chart1" "TestChart" "单位:万元")
      (send xlsx set-line-chart-x-data! "Chart1" "Sheet1" "A1-L1")
      (send xlsx add-line-chart-serial! "Chart1" "Sheet1" "数据1" "A2-L2")
      (send xlsx add-line-chart-serial! "Chart1" "Sheet1" "数据2" "A3-L3")

      (send xlsx add-line-chart-sheet "Chart2" "TestChart" "单位:万元")
      (send xlsx set-line-chart-x-data! "Chart2" "Sheet1" "A1-L1")
      (send xlsx add-line-chart-serial! "Chart2" "Sheet1" "数据1" "A2-L2")
      (send xlsx add-line-chart-serial! "Chart2" "Sheet1" "数据2" "A3-L3")

      (write-xlsx-file xlsx "test1.xlsx")
      ))
   ))

(run-tests write-test)
