# -*- coding: utf-8 -*-

from os import listdir
from os.path import isfile, join
import sys
import os
import pdf2image
from pdf2image import convert_from_path, convert_from_bytes
import comtypes.client
import time
from PIL import Image
mypath = "C:\WordDIR"
tiffpath= "C:\TiffDIR\\"
wdFormatPDF = 17
for f in listdir(mypath) :
    if isfile(join(mypath, f)):
        in_file=os.path.abspath(mypath+"\\"+f)
        if(os.path.splitext(f)[1] =='.docx'):
            out_file = os.path.abspath(mypath+'\\temp.pdf')
            word = comtypes.client.CreateObject('Word.Application')
            time.sleep(3)
            doc = word.Documents.Open(in_file)        
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()        
            convert_from_path((mypath+'\\temp.pdf'), dpi=200,  fmt='tiff',output_folder=tiffpath)
            for j in listdir(tiffpath) :
                if isfile(join(tiffpath, j)):
                    base=os.path.basename(tiffpath+j)
                    if(os.path.splitext(base)[1] =='.ppm'):
                        image = Image.open(tiffpath+j)
                        image.save(tiffpath+os.path.splitext(base)[0]+'.tiff')
                        os.rename(tiffpath+os.path.splitext(base)[0]+'.tiff',tiffpath+os.path.splitext(f)[0]+'.tiff') 
                        os.remove(tiffpath+j) 
            os.remove(mypath+'\\temp.pdf') 

