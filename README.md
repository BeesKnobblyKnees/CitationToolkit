# Citation Toolkit
### For pediatric orthopedics manuscript editing — Tachdjian's Pediatric Orthopaedics

A Streamlit web app for medical librarians and editors working with EndNote and Word documents.

## Tools included

| App | Purpose |
|-----|---------|
| **App 1 — Citation Repair** | Find missing citation placeholders and match them to your EndNote library |
| **App 2 — Broken Citation Fixer** | Fix citations not recognized by EndNote after document merging or conversion |
| **App 3 — Reference Comparator** | Compare two reference lists, find missing or extra refs |

## Run locally

```bash
pip install -r requirements.txt
streamlit run app_v2.py
```

## File structure

```
CitationToolkit/
  app_v2.py              <- Main app (run this)
  citation_repair.py     <- App 1 logic
  ref_comparator.py      <- App 3 logic
  requirements.txt       <- Python dependencies
  .streamlit/
    config.toml          <- Streamlit settings
  README.md              <- This file
```
