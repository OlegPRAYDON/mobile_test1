import requests
import lxml.html
from kivy.app import App
from kivy.core.window import Window
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.config import Config
import os
import wget
from PIL import ImageGrab
import win32com.client
from PIL import Image
#===============================

class mainApp(App):
    def build(self):
        #window_size = Window.size
        #print(window_size)
        Config.set('graphics', 'resizable', '0')
        self.table = self.parse("http://www.s10034.edu35.ru/расписание/")
        if not os.path.isdir("data"):
            os.mkdir("data")
        box = BoxLayout(orientation='vertical')
        box_href = BoxLayout(orientation='vertical')
        self.img_path_sp = []
        global i
        for i in range(len(self.table[0])):
            print('Beginning file download with wget module')
            url = self.table[1][i]
            if url[-4:] == "xlsx":
                name = f'shudle_{i}'
                wget.download(url, f"data\{name}.xlsx")
                self.exel2jpg(name)
                image_path = self.img_path_sp[i]
                img = Image.open(f"data\{image_path}")
                # изменяем размер
                #new_image = img.resize((200, 385))
                #new_image.show()
                # сохранение картинки
                #new_image.save(self.img_path_sp[i])
        box.add_widget(box_href)
        return box

    def parse(self, url):
        api = requests.get(url)
        tree = lxml.html.document_fromstring(api.text)
        text = tree.xpath('//*[@id="content"]/div[1]/div//p/a/text()')
        href = tree.xpath('//*[@id="content"]/div[1]/div//p/a/@href')
        return text, href

    def delete(self, name):
        if os.path.isfile(f'data\{name}.xlsx'):
            os.remove(f'data\{name}.xlsx')
            print("success")
        else:
            print("File doesn't exists!")
            
    
    def get_path(self, *path):
        root_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(root_path, *path)
    
    def exel2jpg(self, name):
        xlsx_path = self.get_path("data", f"{name}.xlsx")
        # https://stackoverflow.com/questions/44850522/python-export-excel-sheet-range-as-image
        # https://docs.microsoft.com/ru-ru/office/vba/api/excel.application(object)
        client = win32com.client.Dispatch("Excel.Application")
        # https://docs.microsoft.com/ru-ru/office/vba/api/excel.workbooks
        wb = client.Workbooks.Open(xlsx_path)
        # https://docs.microsoft.com/ru-ru/office/vba/api/excel.worksheet
        # ws = wb.ActiveSheet
        wsheets = wb.Worksheets.Count
        for i in range(1, wsheets):
            ws = wb.Worksheets(i)
            # for v in ws.Range("A1"): print(v) 
            # https://docs.microsoft.com/ru-ru/office/vba/api/excel.range.copypicture
            ws.Range("A1:CF42").CopyPicture(Format = 2)
            img = ImageGrab.grabclipboard()
            img_path = f"{name}_{i}.jpg"
            self.img_path_sp.append(img_path)
            img.save(self.get_path('data', img_path))
        wb.Close() # иначе табл будет открыта
        client.Quit()
        self.delete(name)     
#=============================
if __name__ == "__main__":
    app = mainApp()
    app.run()
