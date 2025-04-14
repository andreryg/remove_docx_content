import zipfile
import os
from removeAllButOneParagraph import removeAllButOneParagraph

def unzipFile(zip_file_path, output_dir):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

def rezipFile(output_dir, zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'w') as zip_ref:
        for foldername, subfolders, filenames in os.walk(output_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zip_ref.write(file_path, os.path.relpath(file_path, output_dir))

def main(path_to_word_file):

    #Registreringsregler
    unzipFile(path_to_word_file, "unzipped")
    removeAllButOneParagraph("unzipped/word/document.xml", "Registreringsregler med eksempler", "Registreringsregler", "Overskrift2")
    rezipFile("unzipped", "42_registreringsregler.docx")

    #Eksempler
    unzipFile(path_to_word_file, "unzipped")
    removeAllButOneParagraph("unzipped/word/document.xml", "Registreringsregler med eksempler", "Eksempler", "Overskrift2")
    rezipFile("unzipped", "42_eksempler.docx")

if __name__ == "__main__":
    main("41.docx")