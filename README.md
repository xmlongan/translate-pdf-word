# translate_pdf_word package (The project is dead!)

This is a **Python** package for translating pdf or word files into their Chinese versions.

### Install the package

```
pip install translate-pdf-word
```

### Usage

```python
from translate_pdf_word import Word2word
# substitute "your path/to/word.docx" with your word file name or path to it
word = Word2word.Word2word("your path/to/word.docx")
word.translate_all_in_one_shot()
word.save2docx()
```
