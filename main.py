import zipfile
import os
import shutil
from removeAllButOneParagraph import removeAllButOneParagraph # type: ignore

def unzipFile(zip_file_path, output_dir):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

def rezipFile(output_dir, zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'w') as zip_ref:
        for foldername, subfolders, filenames in os.walk(output_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zip_ref.write(file_path, os.path.relpath(file_path, output_dir))

def main(vtid):
    #Registreringsregler
    unzipFile(f"{vtid}.docx", "unzipped")
    removeAllButOneParagraph("unzipped/word/document.xml", "Registreringsregler med eksempler", "Registreringsregler", "Overskrift2")
    rezipFile("unzipped", f"{vtid}_registreringsregler.docx")
    shutil.rmtree("unzipped")

    #Eksempler
    unzipFile(f"{vtid}.docx", "unzipped")
    removeAllButOneParagraph("unzipped/word/document.xml", "Registreringsregler med eksempler", "Eksempler", "Overskrift2")
    rezipFile("unzipped", f"{vtid}_eksempler.docx")
    shutil.rmtree("unzipped")

if __name__ == "__main__":
    vtid = 846
    main(vtid)