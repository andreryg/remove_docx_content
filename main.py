import zipfile
import os
import shutil
import pypandoc
import docx2pdf
from paragraphToDocx import subsubparagraphToDocx, findSubsubparagraphs, subparagraphToDocx # type: ignore
from fixLinks import fixLinks # type: ignore
from docxToHtml import WordOle, WordSaveFormat # type: ignore

def unzipFile(zip_file_path, output_dir):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

def rezipFile(output_dir, zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'w') as zip_ref:
        for foldername, subfolders, filenames in os.walk(output_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zip_ref.write(file_path, os.path.relpath(file_path, output_dir))

def docx2html(docx_file_path):
    pypandoc.convert_file(docx_file_path, 'html', outputfile=docx_file_path.replace(".docx", ".html"), extra_args=['--standalone'])

def main(vtid):
    #Registreringsregler
    unzipFile(f"{vtid}.docx", "unzipped")
    subparagraphToDocx("unzipped/word/document.xml", "Registreringsregler med eksempler", "Overskrift1", "Registreringsregler", "Overskrift2")
    fixLinks("unzipped/word/document.xml")
    rezipFile("unzipped", f"{vtid}_registreringsregler.docx")
    word_ole = WordOle(f"Python Scripts/remove_docx_content/{vtid}_registreringsregler.docx")
    word_ole.show()
    word_ole.save(f"Python Scripts/remove_docx_content/{vtid}_registreringsregler.html", WordSaveFormat.wdFormatHTML)
    word_ole.close()
    shutil.rmtree("unzipped")
    shutil.rmtree(f"{vtid}_registreringsregler-filer")

    #Eksempler
    unzipFile(f"{vtid}.docx", "unzipped")
    subsubparagraphs = findSubsubparagraphs("unzipped/word/document.xml", "Eksempler", 'Overskrift3')
    for i,subparagraph_header in enumerate(subsubparagraphs):
        print(subparagraph_header)
        unzipFile(f"{vtid}.docx", "unzipped")
        subsubparagraphToDocx("unzipped/word/document.xml", subparagraph_header, "Overskrift1", 'Overskrift3')
        fixLinks("unzipped/word/document.xml")
        subparagraph_header = subparagraph_header.replace("/", "£")
        rezipFile("unzipped", f"{vtid}_Eks_{i+1}_{subparagraph_header}.docx")
        docx2pdf.convert(f"{vtid}_Eks_{i+1}_{subparagraph_header}.docx", f"{vtid}_Eks_{i+1}_{subparagraph_header}.pdf")
        shutil.rmtree("unzipped")
    

if __name__ == "__main__":
    vtid = 208
    main(vtid)