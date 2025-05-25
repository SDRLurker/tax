import pandas as pd
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.spinner import Spinner
from kivy.clock import Clock
from kivy.uix.button import Button
from kivy.core.text import LabelBase
from kivy.graphics import Color, Rectangle
import threading
import time
import collect
from kivy.metrics import dp
import openpyxl
from datetime import datetime
import os
import sys
import platform
import subprocess
import shutil
import webbrowser


secu_dic = {
    "한국투자증권": collect.한국투자증권(0),
    "삼성증권": collect.삼성증권(0),
    "키움증권": collect.키움증권(0),
    "신한투자증권": collect.신한투자증권(0),
    "하나증권": collect.하나증권(0)
}


class DataFrameApp(App):
    icon = "icon.png"
    title = "해외 주식 양도세 프로그램(2025.05.25)"
    def build(self):
        return DataFrameBox()


def open_excel(file_path):
    if platform.system() == "Windows":
        os.startfile(file_path)
    elif platform.system() == "Darwin":  # macOS
        subprocess.call(["open", file_path])
    else:  # Linux
        subprocess.call(["xdg-open", file_path])


class DataFrameBox(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.spinners = []

    def update_rect(self, rect):
        def _update(instance, *args):
            rect.pos = instance.pos
            rect.size = instance.size
        return _update

    def on_kv_post(self, base_widget):
        Clock.schedule_once(self.add_custom_row, 0)

    def add_custom_row(self, idx=0):
        grid_layout = self.ids.grid_layout
        spinner = Spinner(
            text="한국투자증권",
            values=["한국투자증권", "삼성증권", "키움증권", "신한투자증권", "하나증권"],
            size_hint_x=0.25,
            height=dp(50)
        )
        spinner.name=f"spinner{idx}"
        grid_layout.add_widget(spinner)
        grid_layout.add_widget(Label(text="", size_hint_x=0.25))
        grid_layout.add_widget(Label(text="", size_hint_x=0.25))
        grid_layout.add_widget(Label(text="", size_hint_x=0.25))
        self.spinners.append(spinner)

    def remove_custom_row(self):
        grid_layout = self.ids.grid_layout
        if grid_layout.children:
            for _ in range(4):
                grid_layout.remove_widget(grid_layout.children[0])
        self.spinners.pop()

    def init_tax(self):
        self.spinners = []
        self.ids.msg.text = ""
        grid_layout = self.ids.grid_layout
        grid_layout.clear_widgets()  # 기존 데이터 삭제
        grid_layout.add_widget(Label(text="증권사", size_hint_x=0.25))
        grid_layout.add_widget(Label(text="손익", size_hint_x=0.25))
        grid_layout.add_widget(Label(text="제비용", size_hint_x=0.25))
        grid_layout.add_widget(Label(text="차감손익", size_hint_x=0.25))
        self.add_custom_row()
        self.ids.init_button.disabled = False
        self.ids.add_button.disabled = False
        self.ids.remove_button.disabled = True
        self.ids.collect_button.disabled = False
        self.ids.excel_button.disabled = True

    def add_tax(self):
        idx = len(self.spinners)
        self.add_custom_row(idx)
        if len(self.spinners) > 1:
            self.ids.remove_button.disabled = False

    def remove_tax(self):
        if len(self.spinners) > 1:
            self.remove_custom_row()
        if len(self.spinners) == 1:
            self.ids.remove_button.disabled = True

    def create_dataframe(self):
        # 비동기적으로 데이터프레임을 가져오도록 스레드 사용
        threading.Thread(target=self.load_dataframe, daemon=True).start()

    def load_dataframe(self):
        try:
            self.ids.init_button.disabled = True
            self.ids.add_button.disabled = True
            self.ids.remove_button.disabled = True
            self.ids.collect_button.disabled = True
            self.tax_list = []

            for i, s in enumerate(self.spinners):
                c = secu_dic[s.text]
                self.tax_list.append((s.text, c))
            all_sum_df = pd.DataFrame(columns=["증권사","손익","제비용","차감손익"])
            Clock.schedule_once(lambda dt: self.update_ui(all_sum_df))
            self.dfs = []
            self.sum_dfs = []
            손익합 = 0
            제비용합 = 0
            차감손익합 = 0
            for name, sec in self.tax_list:
                self.ids.msg.text = f"{name} - 수집 시작!"
                with sec as a:
                    while True:
                        time.sleep(10)
                        if a.is_login():
                            break
                        self.ids.msg.text = f"{name} - 로그인이 필요합니다."
                    self.ids.msg.text = f"{name} - 로그인 성공"
                    a.go_collect_page()
                    while a.is_login():
                        self.ids.msg.text = f"{name} - 비밀번호 입력이 필요합니다."
                        time.sleep(10)
                        if a.is_password():
                            break
                    while True:
                        self.ids.msg.text = f"{name} - 데이터 수집을 시작합니다."
                        df = a.collect()
                        if isinstance(df, pd.DataFrame):
                            break
                        self.ids.msg.text = f"{name} - 데이터 수집 실패. 재시도 합니다."
                        time.sleep(10)
                    self.dfs.append(df)
                    df_sum = df.sum()
                    손익 = df_sum.iloc[10] - df_sum.iloc[13]
                    제비용 = df_sum.iloc[14]
                    차감손익 = 손익-제비용
                    손익합 += 손익
                    제비용합 += 제비용
                    차감손익합 += 차감손익
                    sum_data = [[name, 손익, 제비용, 차감손익]]
                    sum_df = pd.DataFrame(sum_data)
                    self.sum_dfs.append(sum_df)
                    all_sum_df = pd.concat(self.sum_dfs)
                    Clock.schedule_once(lambda dt: self.update_ui(all_sum_df))
            sum_data = [['합계', 손익합, 제비용합, 차감손익합]]
            sum_df = pd.DataFrame(sum_data)
            self.sum_dfs.append(sum_df)
            result_data = [['양도소득금액','과세표준','산출세액','지방세']]
            과세표준 = 차감손익합 - 2500000 if 차감손익합 >= 2500000 else 0
            산출세액 = 과세표준//5
            지방세 = 산출세액//10
            result_data.append([차감손익합, 과세표준, 산출세액, 지방세])
            result_df = pd.DataFrame(result_data)
            self.sum_dfs.append(result_df)
            all_sum_df = pd.concat(self.sum_dfs)
            Clock.schedule_once(lambda dt: self.update_ui(all_sum_df))
            self.ids.msg.text = "수집 완료!"
            self.ids.init_button.disabled = False
            self.ids.excel_button.disabled = False
        except Exception as e:
            self.ids.msg.text = f"오류 발생!"
            print(f"An error occurred: {e}")
            Clock.schedule_once(lambda dt: self.init_tax())

    def excel(self):
        self.ids.excel_button.disabled = True
        self.ids.msg.text = "엑셀을 시작합니다."
        time.sleep(1)
        df_tax = pd.concat(self.dfs, ignore_index=True)
        file_path = '주식_엑셀업로드_양식.xlsx'
        wb = openpyxl.load_workbook(file_path)
        sheet_name = '자료'
        ws = wb[sheet_name]
        for i, row in df_tax.iterrows():
            for j, value in enumerate(row, start=1):
                ws.cell(row=2+i, column=j, value=value)
        now = datetime.now()
        file_path = '주식_엑셀업로드_양식_%s.xlsx' % now.strftime('%Y%m%d')
        wb.save(file_path)
        try:
            open_excel(file_path)
            self.ids.msg.text = "엑셀을 실행했습니다."
        except Exception as e:
            self.ids.msg.text = "엑셀 파일을 여는 중 오류 발생"
            print(e)
        finally:
            self.ids.excel_button.disabled = False

    def update_ui(self, df):
        # GridLayout 가져오기
        grid_layout = self.ids.grid_layout
        grid_layout.clear_widgets()  # 기존 데이터 삭제
        grid_layout.add_widget(Label(text='증권사', bold=True))
        grid_layout.add_widget(Label(text='손익', bold=True))
        grid_layout.add_widget(Label(text='제비용', bold=True))
        grid_layout.add_widget(Label(text='차감손익', bold=True))
        # 데이터프레임의 데이터를 추가
        for index, row in df.iterrows():
            for col in range(4):
                label = Label(text=str(row[col]))
                grid_layout.add_widget(label)
        # 앞으로 진행할 데이터를 표시
        for idx, spinner in enumerate(self.spinners):
            if idx < len(df):
                continue
            elif idx >= len(df):
                for col in range(4):
                    s = spinner.text if col == 0 else ""
                    # 현재 진행 중인 행을 표시 연두색 (0.91, 0.96, 0.91)
                    if idx == len(df):
                        label = Label(text=s, color=(0, 0, 0, 1))
                        with label.canvas.before:
                            Color(0.91, 0.96, 0.91, 1)
                            rect = Rectangle(pos=label.pos, size=label.size)
                        label.bind(pos=self.update_rect(rect), size=self.update_rect(rect))
                    else:
                        label = Label(text=s)
                    grid_layout.add_widget(label)

    def homepage(self):
        webbrowser.open("https://sdrlurker.notion.site/166d1dff257980d3a8e8c71e5c022e81")


def resource_path(relative_path):
    """PyInstaller 사용 시 리소스 파일 경로 반환"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def copy_file_to_current_directory(file_name):
    file_path = resource_path(file_name)
    destination_path = os.path.join(os.getcwd(), file_name)
    if not os.path.exists(destination_path):
        try:
            shutil.copy(file_path, destination_path)
            print(f"Font file copied to: {destination_path}")
        except Exception as e:
            print(f"Error copying font file: {e}")
            sys.exit(1)
    else:
        print(f"Font file already exists at: {destination_path}")
    return destination_path


if __name__ == "__main__":
    font_path = copy_file_to_current_directory("NanumGothic.otf")
    excel_path = copy_file_to_current_directory("주식_엑셀업로드_양식.xlsx")
    icon_path = copy_file_to_current_directory("icon.png")
    DataFrameApp().run()
