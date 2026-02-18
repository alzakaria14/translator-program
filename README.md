# 📄 DOCX Translator — LibreTranslate

A Python script to automatically translate Microsoft Word (`.docx`) documents from **Indonesian to English** (or any language pair) using a self-hosted [LibreTranslate](https://github.com/LibreTranslate/LibreTranslate) instance.

## ✨ Features

- Translates all paragraphs in a `.docx` file, including text inside **tables**
- **Batch processing** to maximize throughput while keeping requests stable
- Automatically **splits oversized paragraphs** that exceed the batch character limit
- **Retry logic** with exponential backoff on failed requests
- Falls back to original text if translation completely fails (no data loss)
- Progress bar via `tqdm`
- Preserves paragraph **styles** (Heading, Normal, etc.)

## 📋 Requirements

- Python 3.8+
- A running [LibreTranslate](https://github.com/LibreTranslate/LibreTranslate) instance (default: `http://localhost:5009`)

### Python dependencies

```bash
pip install python-docx requests tqdm
```

## 🚀 Usage

1. **Start your LibreTranslate server** (e.g. via Docker):

```bash
docker run -ti --rm -p 5009:5000 libretranslate/libretranslate
```

1. **Place your Word document** in the project folder and edit the last lines of `translate_docx.py`:

```python
translate_docx("input.docx", "output_en.docx")
```

1. **Run the script**:

```bash
python translate_docx.py
```

The translated document will be saved as `output_en.docx`.

## ⚙️ Configuration

All configuration is at the top of `translate_docx.py`:

| Variable | Default | Description |
| --- | --- | --- |
| `LIBRETRANSLATE_URL` | `http://localhost:5009/translate` | LibreTranslate endpoint |
| `SOURCE_LANG` | `id` | Source language code |
| `TARGET_LANG` | `en` | Target language code |
| `MAX_TOTAL_CHARS_PER_BATCH` | `20000` | Max total characters per API request |
| `MAX_ITEMS_PER_BATCH` | `50` | Max paragraphs per API request |
| `RETRY_LIMIT` | `4` | Number of retry attempts on failure |
| `REQUEST_TIMEOUT` | `180` | HTTP request timeout (seconds) |

You can increase `MAX_TOTAL_CHARS_PER_BATCH` and `MAX_ITEMS_PER_BATCH` if your machine can handle heavier loads.

## ⚠️ Limitations

- Inline formatting (bold, italic, font size, etc.) inside paragraphs is **not preserved** — only paragraph-level styles (e.g. Heading 1, Normal) are kept.
- Requires a self-hosted LibreTranslate instance; the script does not call any external cloud API.
- Translation quality depends on your LibreTranslate language models.

## 📂 Project Structure

```text
.
├── translate_docx.py   # Main script
├── input.docx          # Your source document (not included)
├── output_en.docx      # Translated output (generated)
└── README.md
```

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

## 📜 License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
