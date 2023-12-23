import pandas as pd
import numpy as np
import time
import os
import sys
import barcode
from reportlab.graphics import renderPDF
from reportlab.pdfgen import canvas
from svglib.svglib import svg2rlg
from reportlab.lib.units import mm
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
import fitz
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from tkinter.filedialog import askopenfilename, askdirectory
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk


base_dir = os.path.dirname(__file__)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
        #base_path = base_dir
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

pdfmetrics.registerFont(TTFont('Arial Narrow', resource_path('src/fonts/Arial Narrow.ttf')))
pdfmetrics.registerFont(TTFont('Arial Narrow Bold', resource_path('src/fonts/Arial Narrow Bold.ttf')))
pdfmetrics.registerFont(TTFont('Courier New', resource_path('src/fonts/Courier New.ttf')))
pdfmetrics.registerFont(TTFont('Courier New Bold', resource_path('src/fonts/Courier New Bold.ttf')))
pdfmetrics.registerFont(TTFont('GT Thin', resource_path('src/fonts/GT-America-Standard-Thin.ttf')))
registerFontFamily('Courier New',normal='Courier New',bold='Courier New Bold')

def m2px(mms):
    return(int((mms/25.4)*72))

def generate_barcode(EAN13):
    EAN = barcode.get_barcode_class('ean13')
    my_ean = EAN(EAN13, guardbar=True)
    ean_path = label_output.get()
    my_ean.save(os.path.join(ean_path,"EAN13")+os.sep+"EAN13"+EAN13)
    return os.path.join(ean_path,"EAN13")+os.sep+"EAN13"+EAN13

def coord(x, y, height, unit=mm):
    x, y = x * unit, height -  y * unit
    return x, y

def scale(drawing, scaling_factor):
    """
    Scale a reportlab.graphics.shapes.Drawing()
    object while maintaining the aspect ratio
    """
    scaling_x = scaling_factor
    scaling_y = scaling_factor

    drawing.width = drawing.minWidth() * scaling_x
    drawing.height = drawing.height * scaling_y
    drawing.scale(scaling_x, scaling_y)
    return drawing

def paste_pdf(image, base):
    base_pdf = PdfReader(base, decompress=False).pages
    base_pdf = pagexobj(base_pdf[0])
    image.doForm(makerl(image, base_pdf))
    image.showPage()
    image.save()

def add_image(image_path, base_pdf, scaling_factor, w, h, output_name, x, y, output_path):
    w, h = (m2px(w), m2px(h))
    my_canvas = canvas.Canvas(output_path+output_name+'.pdf', pagesize=(w, h))
    drawing = svg2rlg(image_path)
    scaled_drawing = scale(drawing, scaling_factor)
    renderPDF.draw(scaled_drawing, my_canvas, *coord(m2px(x), m2px(y), h, mm))
    #my_canvas.drawString(*coord(0, 10, h, mm), 'My SVG Image')
    paste_pdf(my_canvas, base_pdf)

def change_text(input_pdf, variable, text_change, fontsize=6, fontname='Courier'):
    doc = fitz.open(input_pdf)  # the file with the text you want to change
    search_term = variable
    for page in doc:
        found = page.search_for(search_term)  # list of rectangles where to replace
        for item in found:
            page.add_redact_annot(item, '')  # create redaction for text
            page.apply_redactions()  # apply the redaction now
            page.insert_text(item.bl - (0, 3), text_change, fontsize=fontsize, fontname=fontname)
            # https://documentation.help/PyMuPDF/shape.html#Shape.insertText
    doc.save(input_pdf, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)

def sams_inner_label_2(sku, style_name, product_type, colour, jde, base_pdf, output_path, ean_path):
    add_image(ean_path, base_pdf, 0.7, 75, 30, 'Inner_Label_2_'+sku, 10, 10.5, output_path)
    change_text(output_path+'Inner_Label_2_'+sku+'.pdf', "{{style_name}}", style_name, fontsize=6, fontname='Courier-Bold')
    change_text(output_path+'Inner_Label_2_'+sku+'.pdf', "{{product_type}}", product_type, fontsize=6, fontname='Courier')
    change_text(output_path+'Inner_Label_2_'+sku+'.pdf', "{{colour}}", colour, fontsize=6, fontname='Courier-Bold')
    change_text(output_path+'Inner_Label_2_'+sku+'.pdf', "{{jde}}", jde)
    sku2 = sku[:-3]+'  '+'U'
    change_text(output_path+'Inner_Label_2_'+sku+'.pdf', "{{sku}}", sku2)

def sams_inner_label_front(sku, output_path):
    pdf = fitz.open(resource_path('src/Inner_Label_Front.pdf'))
    pdf.save(output_path+'Inner_Label_Front_'+sku+'.pdf')

def sams_inner_label_rear(sku, ext_mat, int_mat, country_orig, base_pdf, output_path):
    w, h = (30*mm, 75*mm)
    my_canvas = canvas.Canvas(output_path+'Inner_Label_Rear_'+sku+'.pdf', pagesize=(w, h))
    my_Style = ParagraphStyle('materials',
                              fontName='Arial Narrow Bold',
                              #backColor='#F1F1F1',
                              fontSize=4.5,
                              #borderColor='#FFFF00',
                              borderWidth=0,
                              borderPadding=(2,2,2),
                              leading=4,
                              alignment=1
                              )

    text_1 = 'EXTERIOR: '+ext_mat+'<BR/>'
    text_2 = 'FORRO: '+int_mat
    p1 = Paragraph(text_1+text_2, my_Style)
    p1.wrapOn(my_canvas,23*mm, 6*mm)
    p1.drawOn(my_canvas, *coord(3.45, 55.5, h, mm))

    my_Style_2 = ParagraphStyle('materials',
                              fontName='Arial Narrow',
                              #backColor='#F1F1F1',
                              fontSize=4.5,
                              #borderColor='#FFFF00',
                              borderWidth=0,
                              borderPadding=(2,2,2),
                              leading=4,
                              alignment=1
                              )
    text_3 = 'HECHO EN: '+country_orig+'<BR/>'
    text_4 = 'FEITO NA: '+country_orig
    p1 = Paragraph(text_3+text_4, my_Style_2)
    p1.wrapOn(my_canvas, 11.3*mm, 3.75*mm)
    p1.drawOn(my_canvas, *coord(9.5, 61, h, mm))
    paste_pdf(my_canvas, base_pdf)

def sams_hang_sticker_latam_all(sku,style_name,product_type,colour,jde,eu_sku,base_pdf,output_path,ean_path):
    w, h = (70*mm, 32*mm)
    my_canvas = canvas.Canvas(output_path+'Hangtag_Sticker_Latam_All_'+sku+'.pdf', pagesize=(w, h))
    style_1 = ParagraphStyle('', fontName='Courier New',
                             fontSize=8.27,
                             borderPadding=(2,2,2),
                             leading=7,
                             alignment=0
                            )
    text = '<b>'+style_name.upper()+'</b>'+'<BR/>'+product_type.upper()+'<BR/>'+colour.upper()
    p1 = Paragraph(text, style_1)
    p1.wrapOn(my_canvas,40*mm, 21*mm)
    p1.drawOn(my_canvas, *coord(3.2, 14, h, mm))

    style_2 = ParagraphStyle('', fontName='Courier New',
                             fontSize=6,
                             borderPadding=(2,2,2),
                             leading=7,
                             alignment=0
                            )

    sku2 = sku[:-3]+'  '+'U'
    text2 = '<b>'+jde+'<BR/>'+sku2+'<BR/>'+eu_sku+'</b>'
    p1 = Paragraph(text2, style_2)
    p1.wrapOn(my_canvas,40*mm, 21*mm)
    p1.drawOn(my_canvas, *coord(3.2, 25, h, mm))
    # Barcode
    drawing = svg2rlg(ean_path)
    scaled_drawing = scale(drawing, scaling_factor=0.5)
    renderPDF.draw(scaled_drawing, my_canvas, *coord(40, 28, h, mm))
    paste_pdf(my_canvas, base_pdf)

def sams_hang_sticker_latam_br(sku,style_name,product_type,colour,jde,eu_sku,
                               net_weight, gross_weight, width, liters, height, length, country_orig,
                               base_pdf,output_path,ean_path):
    w, h = (70*mm, 32*mm)
    my_canvas = canvas.Canvas(output_path+'Hangtag_Sticker_Latam_BR_'+sku+'.pdf', pagesize=(w, h))
    style_1 = ParagraphStyle('', fontName='Courier New',
                             fontSize=6,
                             borderPadding=(2,2,2),
                             leading=7,
                             alignment=0
                            )
    text = '<b>'+style_name.upper()+'<BR/>'+product_type.upper()+'<BR/>'+colour.upper()+'</b>'
    p1 = Paragraph(text, style_1)
    p1.wrapOn(my_canvas,40*mm, 21*mm)
    p1.drawOn(my_canvas, *coord(3.2, 14, h, mm))

    style_2 = ParagraphStyle('', fontName='Courier New',
                             fontSize=5,
                             borderPadding=(2,2,2),
                             leading=5,
                             alignment=0
                            )
    text = liters+' L<BR/>'+width+' X '+height+' X '+length+' cm<BR/>'+'PESO NETO: '+net_weight+' kg<BR/>'+'PESO BRUTO: '+gross_weight
    p2 = Paragraph(text, style_2)
    p2.wrapOn(my_canvas,40*mm, 21*mm)
    p2.drawOn(my_canvas, *coord(3.2, 22, h, mm))

    style_3 = ParagraphStyle('', fontName='Courier New',
                             fontSize=5.5,
                             borderPadding=(2,2,2),
                             leading=5,
                             alignment=0
                            )
    sku2 = sku[:-3]+'  '+'U'
    text = '<b>'+jde+' / '+sku2+'<BR/>'+eu_sku.replace("*",'0')+'</b>'
    p3 = Paragraph(text, style_3)
    p3.wrapOn(my_canvas,40*mm, 21*mm)
    p3.drawOn(my_canvas, *coord(3.2, 27, h, mm))

    style_4 = ParagraphStyle('', fontName='GT Thin',
                             fontSize=3.6,
                             borderPadding=(2,2,2),
                             leading=5,
                             alignment=0
                            )
    sku2 = sku[:-3]+'  '+'U'
    text = 'Produzido na '+country_orig
    p4 = Paragraph(text, style_4)
    p4.wrapOn(my_canvas,40*mm, 21*mm)
    p4.drawOn(my_canvas, *coord(42.2, 27, h, mm))

    # Barcode
    drawing = svg2rlg(ean_path)
    scaled_drawing = scale(drawing, scaling_factor=0.5)
    renderPDF.draw(scaled_drawing, my_canvas, *coord(40, 14, h, mm))
    paste_pdf(my_canvas, base_pdf)

def sams_hang_sticker_eu_all(sku,style_name,product_type,colour,jde,eu_sku,base_pdf,output_path,ean_path):
    w, h = (63*mm, 23*mm)
    my_canvas = canvas.Canvas(output_path+'Hangtag_Sticker_EU_All_'+sku+'.pdf', pagesize=(w, h))
    style_1 = ParagraphStyle('', fontName='Courier New',
                             fontSize=6.48,
                             borderPadding=(2,2,2),
                             leading=7,
                             alignment=0
                            )
    text = '<b>'+style_name.upper()+'</b>'+'<BR/>'+product_type.upper()+'<BR/>'+colour.upper()
    p1 = Paragraph(text, style_1)
    p1.wrapOn(my_canvas,40*mm, 21*mm)
    p1.drawOn(my_canvas, *coord(2, 12, h, mm))

    style_2 = ParagraphStyle('', fontName='Courier New',
                             fontSize=6,
                             borderPadding=(2,2,2),
                             leading=7,
                             alignment=0
                            )

    sku2 = sku[:-3]+'  '+'U'
    text2 = '<b>'+jde+'<BR/>'+sku2+'<BR/>'+eu_sku+'</b>'
    p1 = Paragraph(text2, style_2)
    p1.wrapOn(my_canvas,40*mm, 21*mm)
    p1.drawOn(my_canvas, *coord(2, 20, h, mm))
    # Barcode
    drawing = svg2rlg(ean_path)
    scaled_drawing = scale(drawing, scaling_factor=0.5)
    renderPDF.draw(scaled_drawing, my_canvas, *coord(40, 20, h, mm))
    paste_pdf(my_canvas, base_pdf)

def sams_hang_sticker_eu_br(sku,style_name,product_type,colour,jde,eu_sku,
                               net_weight, gross_weight, liters, width, height, length, country_orig,
                               base_pdf,output_path,ean_path):
    w, h = (63*mm, 23*mm)
    my_canvas = canvas.Canvas(output_path+'Hangtag_Sticker_EU_BR_'+sku+'.pdf', pagesize=(w, h))
    style_1 = ParagraphStyle('', fontName='Courier New',
                             fontSize=6,
                             borderPadding=(2,2,2),
                             leading=5,
                             alignment=0
                            )
    text = '<b>'+style_name.upper()+'<BR/>'+product_type.upper()+'<BR/>'+colour.upper()+'</b>'
    p1 = Paragraph(text, style_1)
    p1.wrapOn(my_canvas,40*mm, 21*mm)
    p1.drawOn(my_canvas, *coord(2, 10, h, mm))

    style_2 = ParagraphStyle('', fontName='Courier New',
                             fontSize=5,
                             borderPadding=(2,2,2),
                             leading=4.2,
                             alignment=0
                            )
    text = liters+' L<BR/>'+width+' X '+height+' X '+length+' cm<BR/>'+'PESO NETO: '+net_weight+' kg<BR/>'+'PESO BRUTO: '+gross_weight
    p2 = Paragraph(text, style_2)
    p2.wrapOn(my_canvas,40*mm, 21*mm)
    p2.drawOn(my_canvas, *coord(2, 16.8, h, mm))

    style_3 = ParagraphStyle('', fontName='Courier New',
                             fontSize=4,
                             borderPadding=(2,2,2),
                             leading=5,
                             alignment=0
                            )
    sku2 = sku[:-3]+'  '+'U'
    text = '<b>'+jde+' / '+sku2+'<BR/>'+eu_sku.replace("*",'0')+'</b>'
    p3 = Paragraph(text, style_3)
    p3.wrapOn(my_canvas,40*mm, 21*mm)
    p3.drawOn(my_canvas, *coord(2, 21.8, h, mm))

    style_4 = ParagraphStyle('', fontName='GT Thin',
                             fontSize=3.4,
                             borderPadding=(2,2,2),
                             leading=5,
                             alignment=0
                            )
    sku2 = sku[:-3]+'  '+'U'
    text = 'Produzido na '+country_orig
    p4 = Paragraph(text, style_4)
    p4.wrapOn(my_canvas,40*mm, 21*mm)
    p4.drawOn(my_canvas, *coord(30.3, 22.2, h, mm))

    # Barcode
    drawing = svg2rlg(ean_path)
    scaled_drawing = scale(drawing, scaling_factor=0.45)
    renderPDF.draw(scaled_drawing, my_canvas, *coord(29, 12, h, mm))
    paste_pdf(my_canvas, base_pdf)

def main_hangtag_rear(sku, warranty, output_path):
    input_path = os.path.join(base_dir, './src/Main_Hangtag_Rear_'+warranty+'.pdf')
    pdf = fitz.open(input_path)
    pdf.save(output_path+'Main_Hangtag_Rear_'+warranty+'_'+sku+'.pdf')

def main_hangtag_front(sku, output_path):
    input_path = os.path.join(base_dir, './src/Main_Hangtag_Front.pdf')
    pdf = fitz.open(input_path)
    pdf.save(output_path+'Main_Hangtag_Front_'+sku+'.pdf')

def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = askopenfilename(initialdir="/",
                               title="Select Excel Collaterals Matrix",
                               filetypes=[("Excel files", ".xlsx .xls")]
                               )
    filename = os.path.normpath(filename)
    label_file_entry.delete(0,tk.END)
    label_file_entry.insert(0,filename)
    return None

def Load_output():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = askdirectory()
    filename = os.path.normpath(filename)
    label_output.delete(0,tk.END)
    label_output.insert(0,filename)
    return None

def main(matrix_path, files_output):
    try:
        start_time = time.time()
        process_output.config(foreground="#FFFFFF")
        process_output.config(text='Process Started')

        #path_matrix = os.path.join(base_dir, './src/COLLATERAL INFO MATRIX.xlsx')
        df = pd.read_excel(matrix_path, sheet_name='Datos generales', header=4)
        df = df.fillna(0)
        df['EAN CODE'] = df['EAN CODE'].astype(np.int64)
        df = df.astype({'EAN CODE':'string', 'COUNTRY OF ORIGIN':'string'})
        df['EAN CODE'] = df['EAN CODE'].replace(".0", "")

        for ind in df.index:
            if df['BRAND'][ind] == 'Samsonite' and df['Export?'][ind] == 'Yes':
                sku = df['SKU'][ind]
                brand = df['BRAND'][ind]
                ean13 = df['EAN CODE'][ind]
                eu_sku = df['EU SKU CODE'][ind]
                style_name = df['STYLE NAME'][ind]
                product_type = df['PRODUCT TYPE'][ind]
                liters = str(df['LITRES (L)'][ind])
                net_weight = str(df['NET WEIGHT (KG)'][ind])
                gross_weight = str(df['GROSS WEIGHT (KG)'][ind])
                width = str(df['WIDTH (CM)'][ind])
                length = str(df['LENGTH (CM)'][ind])
                height = str(df['HEIGHT (CM)'][ind])
                warranty = str(df['WARRANTY'][ind])
                colour = df['COLOUR'][ind]
                jde = df['JDE CODE'][ind]
                ext_mat = df['EXTERIOR MATERIAL COMPOSITION (%)'][ind]
                int_mat = df['INTERIOR MATERIAL COMPOSITION (%)'][ind]
                country_orig = df['COUNTRY OF ORIGIN'][ind]

                # Path to output files
                # Add new for every label
                os.path.join(files_output,style_name)+os.sep
                output_path = os.path.join(files_output,style_name)+os.sep
                output_path_2 = os.path.join(files_output,style_name,"Main_Hangtag")
                output_path_3 = os.path.join(files_output,style_name,"Inner_Label_2")
                output_path_4 = os.path.join(files_output,style_name,"Inner_Label")
                output_path_5 = os.path.join(files_output,style_name,"HangSticker_Latam")
                output_path_6 = os.path.join(files_output,style_name,"HangSticker_EU")
                output_path_7 = os.path.join(files_output, "EAN13")
                paths = [output_path, output_path_2, output_path_3, output_path_4,
                        output_path_5, output_path_6, output_path_7]

                for path in paths:
                    if not os.path.exists(path):
                        os.makedirs(path)

                generate_barcode(ean13)

                main_hangtag_rear(sku, warranty, output_path=output_path+os.sep+'Main_HangTag'+os.sep)
                main_hangtag_front(sku, output_path=output_path+'Main_HangTag'+os.sep)
                sams_inner_label_2(sku, style_name, product_type, colour, jde,
                            base_pdf= resource_path('src/Inner_Label_2.pdf'),
                            output_path=output_path+'Inner_Label_2'+os.sep,
                            ean_path=output_path_7+os.sep+"EAN13"+ean13+".svg")
                sams_inner_label_rear(sku, ext_mat, int_mat, country_orig,
                                base_pdf= resource_path('src/Inner_Label_Rear.pdf'),
                                output_path=output_path+'Inner_Label/')
                sams_inner_label_front(sku, output_path=output_path+'Inner_Label/')
                sams_hang_sticker_latam_all(sku, style_name, product_type, colour, jde, eu_sku,
                                base_pdf= resource_path('src/Hang_Sticker_Latam_All.pdf'),
                                output_path=output_path+'HangSticker_Latam/',
                                ean_path=output_path_7+os.sep+"EAN13"+ean13+".svg")
                sams_hang_sticker_latam_br(sku, style_name, product_type, colour, jde, eu_sku,
                                        net_weight, gross_weight, liters, width, length, height, country_orig,
                                        base_pdf= resource_path('src/Hang_Sticker_Latam_BR.pdf'),
                                        output_path=output_path+'HangSticker_Latam/',
                                        ean_path=output_path_7+os.sep+"EAN13"+ean13+".svg")
                sams_hang_sticker_eu_all(sku, style_name, product_type, colour, jde, eu_sku,
                                base_pdf= resource_path('src/Hang_Sticker_EU_All.pdf'),
                                output_path=output_path+'HangSticker_EU/',
                                ean_path=output_path_7+os.sep+"EAN13"+ean13+".svg")
                sams_hang_sticker_eu_br(sku, style_name, product_type, colour, jde, eu_sku,
                                        net_weight, gross_weight, liters, width, length, height, country_orig,
                                        base_pdf= resource_path('src/Hang_Sticker_EU_BR.pdf'),
                                        output_path=output_path+'HangSticker_EU/',
                                        ean_path=output_path_7+os.sep+"EAN13"+ean13+".svg")

        end_time = (time.time() - start_time)
        process_output.configure(text=f'Process finished in {round(end_time,2)}s')
    except:
        process_output.configure(text='There was an error. Try again')

if __name__ == '__main__':
    app = tk.Tk()
    image_logo_ico_path =  resource_path("src/app_icon/SamTag.ico")
    #app.after(250, lambda: app.iconbitmap(image_logo_ico_path))
    app.style = ttk.Style()
    app.style.theme_use('clam')
    app.style.configure('TButton', background = '#065464', foreground = '#d3d3d3', width = 20, borderwidth=0, focusthickness=0)
    app.style.configure("TEntry", fieldbackground="#373737")
    app.style.map('TButton', background=[('active','#85c3cf')])
    image_logo_png_path =  resource_path('src/app_icon/SamTag.png')
    app.title("SamTag")
    app.geometry("500x350")
    app.config(bg = "#212121")

    canvas_tk = tk.Canvas(master=app, width=100, height=100, bg='#212121',bd=0, highlightthickness=0)
    canvas_tk.pack(expand=tk.NO)
    image = Image.open(image_logo_png_path).resize((90,90))
    image = ImageTk.PhotoImage(image)
    canvas_tk.create_image(0, 10, anchor="nw", image=image)

    frame = tk.Frame(master=app, width=1000, height=400, bg='#212121')
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure((0,1), weight=1)
    frame.pack(fill=tk.X)

    matrix_path_browse = ttk.Button(master=frame, text="Browse", command=lambda: File_dialog(), width=10)
    matrix_path_browse.grid(row=1, column=1, padx=10, pady=(30, 0), sticky="e")

    text1 = tk.StringVar()
    text1.set("Select Collateral Matrix   --->    Browse")
    label_file_entry = ttk.Entry(master=frame, width=300, textvariable=text1, foreground='#808080')
    label_file_entry.grid(row=1, column=0, padx=(20,0), pady=(30, 0), sticky="we")

    output_path_browse = ttk.Button(frame, text="Browse", command=lambda: Load_output(), width=10)
    output_path_browse.grid(row=2, column=1, padx=10, pady=(50, 0), sticky="e")

    text2 = tk.StringVar()
    text2.set("Select Output Folder      --->    Browse")
    label_output = ttk.Entry(master=frame, text=text2, foreground="#808080")
    label_output.grid(row=2, column=0, padx=(20,0), pady=(50, 0), sticky="ew")

    generate_frame = tk.Frame(master=app, bg='#212121', pady=30, padx=100)
    generate_frame.pack(fill=tk.X)

    generate = ttk.Button(master=generate_frame, text="Generate Collaterals",
                                       command=lambda: main(label_file_entry.get(), label_output.get()))
    generate.pack(fill=tk.X)

    process_output = tk.Label(master=app, text="Press 'Generate' to start", fg="#808080", background='#212121')
    process_output.pack()

    app.mainloop()