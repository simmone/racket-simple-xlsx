#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/list)
(require rackunit/text-ui)

(require "../../../xlsx/xlsx.rkt")

(require rackunit "chart.rkt")

(define S string-append)

(define (chart-data) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="zh-CN"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="zh-CN" altLang="en-US"/><a:t>测试图表</a:t></a:r><a:endParaRPr lang="en-US" altLang="zh-CN"/></a:p></c:rich></c:tx><c:layout/></c:title><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>金额1</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker><c:cat><c:strRef><c:f>数据页面!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>201601</c:v></c:pt><c:pt idx="1"><c:v>201602</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>数据页面!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>200</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>金额2</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker><c:cat><c:strRef><c:f>数据页面!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>201601</c:v></c:pt><c:pt idx="1"><c:v>201602</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>数据页面!$C$2:$C$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>110</c:v></c:pt><c:pt idx="1"><c:v>210</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:ser><c:idx val="2"/><c:order val="2"/><c:tx><c:v>金额3</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker><c:cat><c:strRef><c:f>数据页面!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>201601</c:v></c:pt><c:pt idx="1"><c:v>201602</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>数据页面!$D$2:$D$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1110</c:v></c:pt><c:pt idx="1"><c:v>1210</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:marker val="1"/><c:axId val="76367360"/><c:axId val="76368896"/></c:lineChart><c:catAx><c:axId val="76367360"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="76368896"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="76368896"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorGridlines/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="zh-CN" altLang="en-US"/><a:t>金额</a:t></a:r><a:endParaRPr lang="en-US" altLang="zh-CN"/></a:p></c:rich></c:tx><c:layout/></c:title><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="76367360"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>
})

(define test-chart
  (test-suite
   "test-chart"

   (test-case
    "test-chart"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "数据页面"
            #:sheet_data
            '(
              ("月份" "金额1" "金额2" "金额3")
              ("201601" 100 110 1110)
              ("201602" 200 210 1210)
              ))
      (send xlsx add-chart-sheet #:sheet_name "Chart1" #:topic "测试图表" #:x_topic "金额")
      
      (send xlsx set-chart-x-data! #:sheet_name "Chart1" #:data_sheet_name "数据页面" #:data_range "A2-A3")
      (send xlsx add-chart-serial! #:sheet_name "Chart1" #:data_sheet_name "数据页面" #:y_topic "金额1" #:data_range "B2-B3")
      (send xlsx add-chart-serial! #:sheet_name "Chart1" #:data_sheet_name "数据页面" #:y_topic "金额2" #:data_range "C2-C3")
      (send xlsx add-chart-serial! #:sheet_name "Chart1" #:data_sheet_name "数据页面" #:y_topic "金额3" #:data_range "D2-D3")

      (check-equal? (write-chart "Chart1" xlsx) (chart-data))
    ))
   
   ))

(run-tests test-chart)

