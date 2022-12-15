#lang racket

(require rackunit/text-ui rackunit)

(require "../../style/border-style.rkt")

(define test-border-style
  (test-suite
   "test-border-style"

   (test-case
    "test-border-style"

    (let ([border_style (border-style-from-hash-code "f00000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-equal? (BORDER-STYLE-top_color border_style) "F00000")
      (check-equal? (BORDER-STYLE-top_mode border_style) "thin")
      (check-equal? (BORDER-STYLE-bottom_color border_style) "0F0000")
      (check-equal? (BORDER-STYLE-bottom_mode border_style) "thick")
      (check-equal? (BORDER-STYLE-left_color border_style) "00F000")
      (check-equal? (BORDER-STYLE-left_mode border_style) "double")
      (check-equal? (BORDER-STYLE-right_color border_style) "000F00")
      (check-equal? (BORDER-STYLE-right_mode border_style) "dashed")
      (check-equal? (BORDER-STYLE-hash_code border_style) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed"))

    (let ([border_style (border-style-from-hash-code "<p><p><p><p><p><p><p>")])
      (check-equal? (BORDER-STYLE-top_color border_style) "")
      (check-equal? (BORDER-STYLE-top_mode border_style) "")
      (check-equal? (BORDER-STYLE-bottom_color border_style) "")
      (check-equal? (BORDER-STYLE-bottom_mode border_style) "")
      (check-equal? (BORDER-STYLE-left_color border_style) "")
      (check-equal? (BORDER-STYLE-left_mode border_style) "")
      (check-equal? (BORDER-STYLE-right_color border_style) "")
      (check-equal? (BORDER-STYLE-right_mode border_style) "")
      (check-equal? (BORDER-STYLE-hash_code border_style) "<p><p><p><p><p><p><p>"))

    (let ([border_style (border-style-from-hash-code "")])
      (check-false border_style)))

   (test-case
    "test-border-style<=?"

    (let ([border1_style #f]
          [border2_style (border-style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style))
      (check-false (border-style=? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style #f])
      (check-false (border-style<? border1_style border2_style))
      (check-false (border-style=? border1_style border2_style)))

   (let ([border1_style #f]
         [border2_style #f])
      (check-false (border-style<? border1_style border2_style))
      (check-true (border-style=? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-false (border-style<? border1_style border2_style))
      (check-true (border-style=? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "000000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style))
      (check-false (border-style=? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "0f0000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "ff0000<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p>thick<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "<p>thin<p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p><p>0f0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "<p><p>ff0000<p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p><p><p>double<p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "<p><p><p>thick<p>00f000<p>double<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p><p><p><p>00f000<p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "<p><p><p><p>00ff00<p>double<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p><p><p><p><p>double<p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "<p><p><p><p><p>thick<p>000f00<p>dashed")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p><p><p><p><p><p>000f00<p>dashed")]
          [border2_style (border-style-from-hash-code "<p><p><p><p><p><p>000ff0<p>dashed")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p><p><p><p><p><p><p>dashed")]
          [border2_style (border-style-from-hash-code "<p><p><p><p><p><p><p>double")])
      (check-true (border-style<? border1_style border2_style)))

    (let ([border1_style (border-style-from-hash-code "<p><p><p><p><p><p><p>")]
          [border2_style (border-style-from-hash-code "<p><p><p><p><p><p><p>")])
      (check-false (border-style<? border1_style border2_style)))
    )

   ))

(run-tests test-border-style)
