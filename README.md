# About the project
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/AContesini/Convert_PDF_to_DOCX_or_vice-versa/edit/main/README.md) 

Very simply, this is a document converter, which will convert all files from a selected directory and transform them into the selected format, where you can even save them in another folder of your choice.
You can transform the file from PDF to DOCX or DOCX to PDF format, the purpose of this program is to convert large amounts of files at once.

Very simply, this is a document converter, which will convert all files from a selected directory and transform them into the selected format, where you can even save them in another folder of your choice.
You can transform the file from PDF to DOCX or DOCX to PDF format, the purpose of this program is to convert large amounts of files at once.

<br/> 

# Language used
### • Python 

## Libraries

### • pdf2docx
  Used for converting PDF to DOCX.
### • docx2pdf
   Used for converting DOCX to PDF.
### • PySimpleGUI
  Used to create graphical interfaces windows, buttons, text boxes, menus.
### • os
   Used to obtain operating system information with just one line of code.

## See the python code  "main.py".
```python

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
```

#
    
## Menu window
![Modelo Web](https://github.com/AContesini/asstes_img/blob/main/Sem%20t%C3%ADtulo.jpg)

For this window and the elements within it, PySimpleGUI was used. It is a Python library that simplifies the creation of graphical interfaces, making the construction of windows, buttons, text boxes and other elements intuitive.

## have 3 types of menus the first to:
### Select type this conversion
Where we have two radio buttons, when one of them is clicked the other leaves the selection we have PDF to DOCX and DOCX to PDF

### browsers

Just below we have 2 search bars where we can place the folder paths directly in the bar or search through the browser buttons, these bars are the one above 'select the directory of the files that were converted' , to find the item we want to convert and the one below 'Select where files will be saved' as it says, it is used to find the directory where the files will be saved

### To convert

And finally we have the To convert button, which will be used to start converting the files

#
Did you like it? feel free to download, if you want to leave a star ⭐ and follow me I'll follow back ,thanks 🙂.
