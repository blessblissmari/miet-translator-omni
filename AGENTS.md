# miet-translator-pro guidance

- read `README.md` and `HANDSOFF.md` before changing architecture.
- this is a static browser app; avoid adding server relays unless the user explicitly accepts that tradeoff.
- current target concept: input pdf/docx/pptx/zip/rar/7z → render pages → global mimo v2.5 ocr/layout blueprint → per-page mimo v2-omni ocr/translation → pdfjs/docx-preview reconciliation → cheap watchdog → docx/pptx build.
- docx must be treated as a real rendered word document, not just markdown extraction. use `docx-preview` paths for source preview/extraction.
- pptx can be treated as pdf-like visual pages; preserve the rendered slide image when uncertain.
- formulas should end as native office math (omml) in docx; keep latex wrapped in `$...$` / `$$...$$` before `docxBuild.ts`.
- figures/charts/diagrams should be cropped or preserved on a white background. avoid inserting dark UI backgrounds into output documents.
- do not print or commit actual api key values. env/build key names may be referenced, secret values must not.
