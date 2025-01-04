from cx_Freeze import setup, Executable

setup(
    name="tax",
    version="1.0",
    description="",
    options={
        "build_exe": {
            "includes": ["os", "sys", "platform", "subprocess", "shutil",
                "webbrowser", "datetime", "time", "threading", 
                "pandas", "kivy", "openpyxl", "selenium", "webdriver_manager"],
        }
    },
    executables=[Executable("show.py")],
)
