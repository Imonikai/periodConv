import PySimpleGUI as sg
import docx
import os

def convert(from_filename, to_filename):
    document = docx.Document(from_filename)

    for paragraph in document.paragraphs:
        if(paragraph.text != ''):
            for run in paragraph.runs:
                run.text = run.text.replace('、','，')
                run.text = run.text.replace('。','．')
    
    document.save(to_filename)


layout = [
    [sg.Text('変換元'), sg.InputText(), sg.FileBrowse(file_types=(('docx','*.docx'),), key='from')],
    [sg.Text('変換先'), sg.InputText(), sg.FolderBrowse(key='to')],
    [sg.Submit('変換'), sg.Cancel('やめる')],
]

window = sg.Window('docxファイルの句読点をカンマピリオドに変換', layout)
while True:
    event, values = window.read()

    if( values['from'] != '' and values['to'] != '' and event == '変換'):
        from_filename = values['from']
        to_path = values['to']
        basename = os.path.basename(from_filename)
        to_filename = f'{to_path}/periodConv-{basename}'
        convert(from_filename, to_filename)
        break

    if event in [None, 'やめる']:
        break

window.close()