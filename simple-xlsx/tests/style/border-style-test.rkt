#lang racket

(require rackunit/text-ui
         rackunit
         "../../style/border-style.rkt")

(define test-border-style
  (test-suite
   "test-border-style"

   (test-case
    "test-update-border-style"

    (let ([border_style (BORDER-STYLE "000000" "thin" "0F0000" "thick" "00F000" "double" "000F00" "dashed")]
          [new_style1 (BORDER-STYLE "0000FF" #f #f #f #f #f #f #f)]
          [new_style2 (BORDER-STYLE #f "thick" #f #f #f #f #f #f)]
          [new_style3 (BORDER-STYLE #f #f "FF0000" #f #f #f #f #f)]
          [new_style4 (BORDER-STYLE #f #f #f "thin" #f #f #f #f)]
          [new_style5 (BORDER-STYLE #f #f #f #f "00FF00" #f #f #f)]
          [new_style6 (BORDER-STYLE #f #f #f #f #f "dashed" #f #f)]
          [new_style7 (BORDER-STYLE #f #f #f #f #f #f "0000F0" #f)]
          [new_style8 (BORDER-STYLE #f #f #f #f #f #f #f "thin")]
          [new_style9 (BORDER-STYLE "FFFFFF" #f "000000" #f "FFFF00" #f "0000FF" #f)]
          )
      (update-border-style border_style new_style1)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thin" "0F0000" "thick" "00F000" "double" "000F00" "dashed"))

      (update-border-style border_style new_style2)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thick" "0F0000" "thick" "00F000" "double" "000F00" "dashed"))

      (update-border-style border_style new_style3)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thick" "FF0000" "thick" "00F000" "double" "000F00" "dashed"))

      (update-border-style border_style new_style4)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thick" "FF0000" "thin" "00F000" "double" "000F00" "dashed"))

      (update-border-style border_style new_style5)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thick" "FF0000" "thin" "00FF00" "double" "000F00" "dashed"))

      (update-border-style border_style new_style6)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thick" "FF0000" "thin" "00FF00" "dashed" "000F00" "dashed"))

      (update-border-style border_style new_style7)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thick" "FF0000" "thin" "00FF00" "dashed" "0000F0" "dashed"))

      (update-border-style border_style new_style8)
      (check-equal? border_style (BORDER-STYLE "0000FF" "thick" "FF0000" "thin" "00FF00" "dashed" "0000F0" "thin"))

      (update-border-style border_style new_style9)
      (check-equal? border_style (BORDER-STYLE "FFFFFF" "thick" "000000" "thin" "FFFF00" "dashed" "0000FF" "thin"))
      )
    )

   ))

(run-tests test-border-style)
