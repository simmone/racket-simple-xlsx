#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-chart-sheet (-> exact-nonnegative-integer? string?)]
          [write-chart-sheet-file (-> path-string? exact-nonnegative-integer? void?)]
          ))

(define S string-append)

(define (write-chart-sheet chart_sheet_name xlsx) @S{
(let ([chart_sheet (sheet-content (send xlsx get-sheet-by-name chart_sheet_name))]
      [x_data_range (line-chart-sheet-x_data_range chart_sheet)]
      [y_data_range_list (line-chart-sheet-y_data_range_list chart_sheet)]
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="zh-CN"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="zh-CN" altLang="en-US"/><a:t>@|topic|</a:t></a:r><a:r>/a:p></c:rich></c:tx><c:layout/></c:title><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:cat><c:numRef><c:f>@|(data-range-sheet_name x_data_range)|!@|(convert-range (data-range-range_str x_data_range))|</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="@|(range-length (data-range-range_str x_data_range)|"/>@|(with-output-to-string
(lambda ()
(let loop ([loop_list 
<c:pt idx="0"><c:v>201601</c:v></c:pt>
<c:pt idx="1"><c:v>201602</c:v></c:pt>
<c:pt idx="2"><c:v>201603</c:v></c:pt>
<c:pt idx="3"><c:v>201604</c:v></c:pt>
<c:pt idx="4"><c:v>201605</c:v></c:pt>
<c:pt idx="5"><c:v>201606</c:v></c:pt>
<c:pt idx="6"><c:v>201607</c:v></c:pt>
<c:pt idx="7"><c:v>201608</c:v></c:pt>
<c:pt idx="8"><c:v>201609</c:v></c:pt>
<c:pt idx="9"><c:v>201610</c:v></c:pt>
<c:pt idx="10"><c:v>201611</c:v></c:pt>
<c:pt idx="11"><c:v>201612</c:v></c:pt>
<c:pt idx="12"><c:v>201701</c:v></c:pt>
</c:numCache>
</c:numRef>
</c:cat>

<c:ser>
<c:idx val="0"/>
<c:order val="0"/>
<c:tx><c:v>金额1</c:v></c:tx>
<c:marker><c:symbol val="none"/></c:marker>
<c:val>
<c:numRef>
<c:f>数据页面!$B$2:$B$14</c:f>
<c:numCache>
<c:formatCode>General</c:formatCode>
<c:ptCount val="13"/>
<c:pt idx="0"><c:v>100</c:v></c:pt>
<c:pt idx="1"><c:v>200</c:v></c:pt>
<c:pt idx="2"><c:v>300</c:v></c:pt>
<c:pt idx="3"><c:v>400</c:v></c:pt>
<c:pt idx="4"><c:v>500</c:v></c:pt>
<c:pt idx="5"><c:v>600</c:v></c:pt>
<c:pt idx="6"><c:v>500</c:v></c:pt>
<c:pt idx="7"><c:v>400</c:v></c:pt>
<c:pt idx="8"><c:v>300</c:v></c:pt>
<c:pt idx="9"><c:v>300</c:v></c:pt>
<c:pt idx="10"><c:v>100</c:v></c:pt>
<c:pt idx="11"><c:v>100</c:v></c:pt>
<c:pt idx="12"><c:v>500</c:v></c:pt>
</c:numCache>
</c:numRef>
</c:val>
</c:ser>

<c:ser>
<c:idx val="1"/>
<c:order val="1"/>
<c:tx><c:v>金额2</c:v></c:tx>
<c:marker><c:symbol val="none"/></c:marker>
<c:val>
<c:numRef>
<c:f>数据页面!$C$2:$C$14</c:f>
<c:numCache>
<c:formatCode>General</c:formatCode>
<c:ptCount val="13"/>
<c:pt idx="0"><c:v>110</c:v></c:pt>
<c:pt idx="1"><c:v>210</c:v></c:pt>
<c:pt idx="2"><c:v>310</c:v></c:pt>
<c:pt idx="3"><c:v>410</c:v></c:pt>
<c:pt idx="4"><c:v>510</c:v></c:pt>
<c:pt idx="5"><c:v>610</c:v></c:pt>
<c:pt idx="6"><c:v>510</c:v></c:pt>
<c:pt idx="7"><c:v>410</c:v></c:pt>
<c:pt idx="8"><c:v>310</c:v></c:pt>
<c:pt idx="9"><c:v>310</c:v></c:pt>
<c:pt idx="10"><c:v>110</c:v></c:pt>
<c:pt idx="11"><c:v>110</c:v></c:pt>
<c:pt idx="12"><c:v>510</c:v></c:pt>
</c:numCache>
</c:numRef>
</c:val>
</c:ser>

<c:ser>
<c:idx val="2"/>
<c:order val="2"/>
<c:tx><c:v>金额3</c:v></c:tx>
<c:marker><c:symbol val="none"/></c:marker>
<c:val>
<c:numRef>
<c:f>数据页面!$D$2:$D$14</c:f>
<c:numCache>
<c:formatCode>General</c:formatCode>
<c:ptCount val="13"/>
<c:pt idx="0"><c:v>1110</c:v></c:pt>
<c:pt idx="1"><c:v>1210</c:v></c:pt>
<c:pt idx="2"><c:v>1310</c:v></c:pt>
<c:pt idx="3"><c:v>1410</c:v></c:pt>
<c:pt idx="4"><c:v>1510</c:v></c:pt>
<c:pt idx="5"><c:v>1610</c:v></c:pt>
<c:pt idx="6"><c:v>1510</c:v></c:pt>
<c:pt idx="7"><c:v>1410</c:v></c:pt>
<c:pt idx="8"><c:v>1310</c:v></c:pt>
<c:pt idx="9"><c:v>1310</c:v></c:pt>
<c:pt idx="10"><c:v>1110</c:v></c:pt>
<c:pt idx="11"><c:v>1110</c:v></c:pt>
<c:pt idx="12"><c:v>1510</c:v></c:pt>
</c:numCache>
</c:numRef>
</c:val>
</c:ser>

<c:marker val="1"/><c:axId val="76367360"/><c:axId val="76368896"/>

</c:lineChart>

<c:catAx><c:axId val="76367360"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="76368896"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="76368896"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="76367360"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>

</c:plotArea>

<c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/>

</c:chart>

</c:chartSpace>
})

(define (write-chart-sheet-file dir typeSeq)
  (with-output-to-file (build-path dir (format "sheet~a.xml" typeSeq))
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-chart-sheet typeSeq)))))

