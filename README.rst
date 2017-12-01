====================
python-docx-template
====================

This forked version adds Markdown support! Currently supported features are:

* Bullet lists
* Code blocks (single backticks and triple backticks)
* Headings (four levels)
* Bold text
* Italic text

In order to have code blocks and headings work, the docx template currently needs pre-patching so that the required styles exist in it. A tool for patching a docx file with custom styles as well as a styles.xml file that can be used right away are available in the util/ directory.
