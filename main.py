
from pdf2docx import Converter
from docx2pdf import convert
import PySimpleGUI as sg
from os import path, listdir

# Interface elements for PDF to DOCX
pdfToDocx = sg.Radio('PDF to DoCX', 'file_type', default=True, key='pdfToDocx')
# Interface elements for DOCX  to PDF
docxToPdf = sg.Radio('Docx to PDF', 'file_type', key='docxToPdf')

# Interface elements for the input files directory
filesDirectory = [
    [sg.Text('select the directory of the files that were converted')],
    [sg.Input(key='folder'), sg.FolderBrowse()]
]

# Interface elements for the output directory
outputDirectory = [
    [sg.Text('Select where files will be saved')], [
        sg.Input(key='output_folder'), sg.FolderBrowse()]
]

# Interface element for the conversion button
button_to_Conversion = [sg.Button('to convert')]

# Combination of layout elements

layout = [
    [sg.Text('Select the type of Conversion')],
    [pdfToDocx],
    [docxToPdf],
    *filesDirectory,
    *outputDirectory,
    button_to_Conversion

]
# Create the window

window = sg.Window('converter', layout)

while True:
    # Read events and values from the window
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    if event == 'to convert':
        try:
            inputFolder = values['folder']
            outputFolder = values['output_folder']
            arqList = listdir(inputFolder)
            # Convert PDF to DOCX if the corresponding option is selected
            if arqList and inputFolder and outputFolder:
                for arquive in arqList:
                    inputFilePath = path.join(inputFolder, arquive)

                    if values['pdfToDocx'] and arquive.lower().endswith('.pdf'):
                        outputFilePath = path.join(
                            outputFolder, arquive.replace('.pdf', '.docx'))
                        c = Converter(inputFilePath)
                        c.convert(outputFilePath)
                  # Convert DOCX to PDF if the corresponding option is selected
                    elif values['docxToPdf'] and arquive.lower().endswith('.docx'):
                        outputFilePath = path.join(
                            outputFolder, arquive.replace('.docx', '.pdf'))
                        convert(inputFilePath, outputFilePath)

                sg.popup('Conversion completed')
       # Handle the case when a file is not found
       except FileNotFoundError:
            pass


window.close()
