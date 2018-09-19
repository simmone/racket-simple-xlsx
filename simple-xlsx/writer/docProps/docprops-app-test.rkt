#lang racket

(require rackunit/text-ui)
(require rackunit "docprops-app.rkt")

(require "../../xlsx/xlsx.rkt")

(define test-docprops-app
  (test-suite
   "test-docprops-app"

   (test-case
    "test-docprops-app"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "数据页面" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet2" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet3" #:sheet_data '((1)))
      (send xlsx add-chart-sheet #:sheet_name "Chart1" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart4" #:topic "Chart4" #:x_topic "")


    (check-equal? (write-docprops-app (get-field sheets xlsx))
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=\"4\" baseType=\"variant\"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant><vt:variant><vt:lpstr>图表</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=\"5\" baseType=\"lpstr\"><vt:lpstr>数据页面</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr><vt:lpstr>Sheet3</vt:lpstr><vt:lpstr>Chart1</vt:lpstr><vt:lpstr>Chart4</vt:lpstr></vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>"))
    ))
   ))

(run-tests test-docprops-app)
