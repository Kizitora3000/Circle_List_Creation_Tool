import sys

import Excel_Init
import Input_to_Excel
import quit_Excel

from kivy.app import App
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.resources import resource_add_path

"""
ウィンドウサイズの固定化
from kivy.config import Config

Config.set("graphics", "resizable", False)
Config.set("graphics", "width", 760)
Config.set("graphics", "height", 570)
"""

LabelBase.register(DEFAULT_FONT, "ipaexg.ttf")
sm = ScreenManager()

if hasattr(sys, "_MEIPASS"):
    resource_add_path(sys._MEIPASS)


class InputScreen(Screen):
    inputted_string_list = []

    def input_button_clicked(self):
        for key in self.ids:
            self.inputted_string_list += self.ids[key].text

        Input_to_Excel.input_to_excel(self.inputted_string_list, title)

        self.clear_button_clicked()

    def clear_button_clicked(self):
        for key in self.ids:
            self.ids[key].text = ""

    def quit_button_clicked(self):
        quit_Excel.quit_excel()


class LoadScreen(Screen):
    def load_button_clicked(self):
        global title
        title = self.ids["text_title"].text
        sm.current = "input"


class CreateScreen(Screen):
    def create_button_clicked(self):
        global title
        title = self.ids["text_title"].text
        Excel_Init.excel_init(title)
        sm.current = "input"


class TitleScreen(Screen):
    def create_button_clicked(self):
        sm.current = "create"

    def load_button_clicked(self):
        sm.current = "load"


class CircleListCreationToolApp1(App):
    def build(self):
        sm.add_widget(TitleScreen(name="title"))
        sm.add_widget(CreateScreen(name="create"))
        sm.add_widget(InputScreen(name="input"))
        sm.add_widget(LoadScreen(name="load"))
        return sm


if __name__ == "__main__":
    CircleListCreationToolApp1().run()
