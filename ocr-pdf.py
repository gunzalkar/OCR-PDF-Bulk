import os   
import sys
import re
import subprocess
import aspose.words as aw
import pypdfium2 as pdfium
from fpdf import FPDF
import shutil

def deleteocr(file):
        os.remove(os.path.join(sys.argv[2],file))
        print("file removed: ",sys.argv[2],file)

def convertPDF2OCR(file):
    print("Converting PDF to OCR")
    pdf_name = file+'.pdf'
    print("Output file name: ",file)
    file = file+"_test.pdf"
    print("file name: ",file)
    command = 'ocrmypdf '+f'"{file}"'+' '+f'"{pdf_name}"'+' -l hin+eng -O 0 --output-type pdf' 
    print("Process will start now!")
    process = subprocess.Popen(command)
    print("Command: ",command)
    stdout, stderr = process.communicate()
    print(stdout, stderr)
    print("Process Done!")
    deleteocr(file)

def PDF2PNG(filename):
    path=filename[:-4]
    print("folder to save images: ",path)
    pdf = pdfium.PdfDocument(sys.argv[1]+'/'+filename)
    if not os.path.exists(os.path.join(sys.argv[1],path)): 
        os.mkdir(os.path.join(sys.argv[1],path))
    n_pages = len(pdf)
    
    print("pages in pdf: ", n_pages)
    for page_number in range(n_pages):
        page = pdf.get_page(page_number)
        print("Processing Page "+str(page_number+1))
        pil_image = page.render_topil(
            scale=200/72,                       #pdf quality 
            rotation=0,
            crop=(0, 0, 0, 0),
            greyscale=False,
        )
        name= f"{path}_image_{page_number+1}.png"
        source =  os.path.join(sys.argv[1],path,name)
        pil_image.save(source)
    return os.path.join(sys.argv[1],path)

def convertPNG2PDF(fileNames,name,dir_path):
    n = 1
    print("Converting PNG into PDF")
    pdf = FPDF()
    for image in fileNames:
        pdf.add_page()
        pdf.image(image,x=0,y=0,w=210,h=297)
        print("added ",n," image")
        n +=1
    pdf.output(name+'_test.pdf')
    print("Deleting image folder: ",dir_path)
    try:
        shutil.rmtree(dir_path)
        print("image directory deleted!")
        convertPDF2OCR(name)
    except Exception as ep:
        print("Error while deleting image folder ",ep)
    
def convertPDF(file):
    if not os.path.exists(sys.argv[2]):
        os.mkdir(sys.argv[2])
    li=[]
    print("working on: ",file)
    dir_path=PDF2PNG(file)
    for images in os.listdir(dir_path):
        if (images.endswith(".png")):
            li.append(dir_path+'\\'+images)
    li.sort(key=lambda f: int(re.sub('\D', '', f)))
    convertPNG2PDF(li,sys.argv[2]+'\\'+file[:-4],dir_path)

if __name__ == "__main__":
    number_of_argument = len(sys.argv)
    if(number_of_argument == 1): 
        print("Enter Folder Location And Destination Location") #If no argument is passed
    elif(number_of_argument == 2): 
        print("Enter Destination Location") #If only one argument is passed
    else:
        print("Folder location",sys.argv[1])
        print("Destination location", sys.argv[2])
        li_f=[]
        for file in os.listdir(sys.argv[1]):
            if not file.endswith(".pdf"):
                continue
            li_f.append(file)
        print(li_f)
        for f in li_f:
            convertPDF(f)
        print("Done with All the Files!! ")
