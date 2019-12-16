# Generate wordlists from Microsoft Office DOCX or PPTX files

Create unique wordlist or dictionary (one word per line) from all the words in MS Office OpenXML formatted documents (docx and pptx only for now).

```bash
git clone https://github.com/proprietary/msoffice-document-tokenizer.git
cd msoffice-document-tokenizer
```

## Examples

Create wordlist file (one word per line) from DOCX document:

```bash
./main.py doc.docx > wordlist
```

Create wordlist file from all DOCX and PPTX in this directory (zsh and bash):

```bash
./main.py *.(docx|pptx) > wordlist
```
