The demo bundles a static regular instance of Noto Sans, generated from the
[Google Fonts Noto Sans variable TTF](https://github.com/google/fonts/tree/main/ofl/notosans)
at width 100 and weight 400 with fontTools. The font is distributed under the
SIL Open Font License in `OFL.txt`.

Browser package consumers supply their own OpenType font bytes to
`registerPdfFont`; the demo font is not copied by `docxport-copy-assets`.
