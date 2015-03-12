#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-docprops-app (-> list? string?)]
          [write-docprops-app-file (-> path-string? list? void?)]
          ))

(define S string-append)
 
(define (write-docprops-app sheet_name_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>@|(number->string (length sheet_name_list))|</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="@|(number->string (length sheet_name_list))|" baseType="lpstr">@|(with-output-to-string
  (lambda ()
    (for-each
      (lambda (sheet_name)
        (printf "<vt:lpstr>~a</vt:lpstr>" sheet_name))
      sheet_name_list)))|</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>
})

(define (write-docprops-app-file dir sheet_name_list)
  (with-output-to-file (build-path dir "app.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-docprops-app sheet_name_list)))))

