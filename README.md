# Remove Docx Content
Python script for removing a paragraph from a docx-file. 
All paragraphs except the ones specified by either paragraph header text or subparagraph header text are removed. The document needs a table of contents, because I chose to use bookmark names, and the bookmark names of paragraph header text has the prefix "_Toc".

The script unzips the file, removes specific paragraphs, and then rezips into a docx file again.
