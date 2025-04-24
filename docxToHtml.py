import win32com.client
import win32com.client.dynamic

class WordSaveFormat:
    wdFormatNone = None
    wdFormatHTML = 8

class WordOle:
    def __init__(self, filename):
        self.filename = filename
        self.word_app = win32com.client.dynamic.Dispatch("Word.Application")
        self.word_doc = self.word_app.Documents.Open(filename)

    def save(self, new_filename=None, word_save_format=WordSaveFormat.wdFormatNone):
        if new_filename:
            self.filename = new_filename
            self.word_doc.SaveAs(new_filename, word_save_format)
        else:
            self.word_doc.Save()

    def close(self):
        self.word_doc.Close(SaveChanges=0)
        # self.word_app.DoClose( SaveChanges = 0 )
        # self.word_app.Close()
        del self.word_app

    def show(self):
        self.word_app.Visible = 1

    def hide(self):
        self.word_app.Visible = 0

"""word_ole = WordOle("Python Scripts/remove_docx_content/208_registreringsregler.docx")
word_ole.show()
word_ole.save("Python Scripts/remove_docx_content/208_registreringsregler.html", WordSaveFormat.wdFormatHTML)
# word_ole.save( "D:\\TestDoc2.docx", WordSaveFormat.wdFormatNone )
word_ole.close()"""