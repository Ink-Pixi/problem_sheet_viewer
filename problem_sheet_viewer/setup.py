from cx_Freeze import setup, Executable

exe=Executable(
     script='main_report.py',
     base="Win32Gui",
     icon="icon/PS_Icon.ico",
     targetName="Report Viewer.exe",
     shortcutName="Report Viewer",
     shortcutDir="DesktopFolder",
     )
build_exe_options = {'packages':['ui', 'data', 'report_viewer'], 'include_files':['icon/', 'C:\Python34\Lib\site-packages\PyQt5\libEGL.dll'], 
                     'includes':['decimal', 'atexit'], 'excludes':['Tkinter'], 'include_msvcr': True}

bdist_msi_options = {
    'upgrade_code': '{66620F3A-DC3A-11E2-B341-002219E9B01E}',
    'add_to_path': False,
    #'initial_target_dir': r'[Program Files] \%s\%s' % ('test', 'test'),
    # 'includes': ['atexit', 'PySide.QtNetwork'], # <-- this causes error
    }
setup(
     version = "0.1",
     description = "View Problem Sheets submitted through Problem Sheet Tool",
     author = "David Hoy",
     name = "Problem Sheet Viewer",
     options = {'build_exe': build_exe_options, 'bdist_msi': bdist_msi_options},
     executables = [exe]
     )