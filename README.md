# 📄 DOCX Translator

A free, automated Word document translator that converts `.docx` files to **60+ languages** — with no API key required. Powered by Google Translate via [`deep-translator`](https://github.com/nidhaloff/deep-translator).

> Preserves your document's original formatting including headings, tables, headers, footers, bold/italic styles, and text boxes.

---

## ✨ Features

- ✅ **Free** — uses Google Translate with no API key or account needed
- ✅ **Format-preserving** — keeps headings, tables, styles, headers & footers intact
- ✅ **60+ languages** — translate to/from any supported language
- ✅ **Batch mode** — translate an entire folder of `.docx` files in one command
- ✅ **Auto-named output** — output file is named automatically with the target language
- ✅ **Retry logic** — gracefully handles network hiccups

---

## 📋 Requirements

- Python 3.8+
- pip

---

## 🚀 Installation

**1. Clone the repository**

```bash
git clone https://github.com/your-username/docx-translator.git
cd docx-translator
```

**2. Install dependencies**

```bash
pip install python-docx deep-translator
```

That's it — no API keys, no accounts, no setup.

---

## 💻 Usage

### Translate a single file (default: English → Russian)

```bash
python translate_docx.py "my_document.docx"
```

Output: `my_document_russian.docx`

---

### Translate to a specific language

```bash
python translate_docx.py "my_document.docx" --to fr
```

Output: `my_document_french.docx`

---

### Specify both source and target language

```bash
python translate_docx.py "my_document.docx" --from fr --to de
```

---

### Custom output file name

```bash
python translate_docx.py "my_document.docx" "translated_output.docx" --to ar
```

---

### Batch translate an entire folder

```bash
python translate_docx.py --batch "C:\Documents\Reports\" --to es
```

Translates all `.docx` files in the folder to Spanish.

---

### See all supported languages

```bash
python translate_docx.py --languages
```

---

## 🌐 Supported Languages

Run `python translate_docx.py --languages` to see the full list. Some popular ones:

| Code | Language | Code | Language | Code | Language |
|------|----------|------|----------|------|----------|
| `ar` | Arabic | `fr` | French | `ru` | Russian |
| `zh-CN` | Chinese (Simplified) | `de` | German | `es` | Spanish |
| `nl` | Dutch | `hi` | Hindi | `sv` | Swedish |
| `en` | English | `it` | Italian | `tr` | Turkish |
| `tl` | Filipino | `ja` | Japanese | `uk` | Ukrainian |
| `fi` | Finnish | `ko` | Korean | `ur` | Urdu |
| `pt` | Portuguese | `pl` | Polish | `vi` | Vietnamese |

---

## 📁 Project Structure

```
docx-translator/
├── translate_docx.py   # Main script
└── README.md
```

---

## ⚙️ How It Works

1. **Reads** the `.docx` file using `python-docx`
2. **Extracts** text from paragraphs, tables, headers, footers, and text boxes
3. **Translates** each block via Google Translate (free tier, no key needed)
4. **Rebuilds** the document preserving all original styles and formatting
5. **Saves** the translated file alongside the original

---

## ⚠️ Limitations

- **Large documents** may take a few minutes due to free API rate limiting
- **Complex mixed formatting** within a single paragraph (e.g. multiple fonts mid-sentence) defaults to the first run's style
- **Images** are preserved as-is (text inside images is not translated)
- **Internet connection** required — translation happens via Google's servers

---

## 🛠️ Troubleshooting

**`No such file or directory` error**
Make sure `translate_docx.py` is in the same folder as your `.docx` file, or provide the full path.

**Filenames with spaces**
Always wrap filenames in quotes:
```bash
python translate_docx.py "my document with spaces.docx"
```

**Translation fails / retries**
Check your internet connection. The script retries automatically up to 3 times per block.

**Unknown language code**
Run `python translate_docx.py --languages` to see all valid codes.

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 🤝 Contributing

Pull requests are welcome! If you find a bug or want to add a feature, feel free to open an issue or submit a PR.
