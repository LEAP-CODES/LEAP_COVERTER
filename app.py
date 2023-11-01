from flask import Flask,flash, render_template,redirect,url_for,request,send_from_directory,send_file
import docx2pdf
import tempfile
import pythoncom
import os
from pdf2docx import Converter,parse

app = Flask(__name__)
app.config['UPLOAD_FOLDER']='uploads'
app.secret_key = "LEapdf8sdufhbcsjkbh34586"

@app.route('/')
def index():
    return render_template('index.html')
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'doc', 'docx'}

@app.route('/ppttopdf',methods=['POST','GET'])
def ppttopdf():
    from pptx import Presentation
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    if request.method=='POST':
        files = request.files['ppt_file']
        tmpdirc = tempfile.mkdtemp()
        if files.name=='':
            flash("No selected files")
            return render_template('powerpoint to pdf preview.html')
        else:
            file_Path = os.path.join(tmpdirc,files.name)
            files.save(file_Path)
            presentation = Presentation(files)
            pdf_file = files.name.replace(".pptx","pdf")
            c = canvas.Canvas(pdf_file,pagesize=letter)
            for slide in presentation.slides:
                image = slide.get_image()
                c.drawImage(image,0,0,width=letter[0],height=letter[1])
            c.save(file_Path)    

        # print("ppptptptt")
        # return "working"
            return send_from_directory(tmpdirc,os.path.basename(file_Path),as_attachment=True)

    

@app.route('/encryptPdf',methods=['POST','GET'])
def encryptpdf():
    from PyPDF2 import PdfReader,PdfWriter
    if request.method=='POST':
        files = request.files['pdf_file']
        password = request.form['password']
        if files.name=='':
            flash("No selected files")
            return render_template('protectpdf preview.html')
        else:
            tmpdirc = tempfile.mkdtemp()
            file_path = os.path.join(tmpdirc,files.name)
            files.save(file_path)
            reader = PdfReader(file_path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(password,use_128bit=True)
            
            encrypted_file_name = 'encrypted.pdf'
            encrypted_file_path = os.path.join(tmpdirc, encrypted_file_name)
            with open(encrypted_file_path, 'wb') as output_pdf:
                writer.write(output_pdf)
            return send_from_directory(tmpdirc, os.path.basename(encrypted_file_path), as_attachment=True)
        
@app.route('/decryptPdf',methods=['POST','GET'])
def decryptpdf():
    from PyPDF2 import PdfReader,PdfWriter
    if request.method=='POST':
        files = request.files['pdf_file']
        password = request.form['password']
        if files.name=='':
            flash("No selected files")
            return render_template('unlockpdf preview.html')
        else:
            tmpdirc = tempfile.mkdtemp()
            file_path = os.path.join(tmpdirc,files.name)
            files.save(file_path)
            reader = PdfReader(file_path)
            writer = PdfWriter()
            try:
                if reader.is_encrypted:
                    reader.decrypt(password)
                for page in reader.pages:
                    writer.add_page(page)    
                encrypted_file_name = 'decrypted.pdf'
                encrypted_file_path = os.path.join(tmpdirc, encrypted_file_name)
                with open(encrypted_file_path, 'wb') as output_pdf:
                    writer.write(output_pdf)
                return send_from_directory(tmpdirc, os.path.basename(encrypted_file_path), as_attachment=True)
            except Exception as e:
                flash("Wrong Password")
                return render_template("unlockpdf preview.html")


@app.route('/imgIntopdf', methods=['POST','GET'])
def imgIntopdf():
    import img2pdf
    if request.method =='POST':
        files = request.files['imagefile']
        if files.name=='':
            flash("No selected files")
            return render_template('JPG to pdf preview.html')
        else:
            tmpDir = tempfile.mkdtemp()
            file_path=os.path.join(tmpDir, files.name + '.pdf')
            files.save(file_path)
            pdfbytes=img2pdf.convert(file_path)
            file = open(file_path,'wb')
            file.write(pdfbytes)
            file.close
            return send_from_directory(tmpDir, os.path.basename(file_path), as_attachment=True)
        
@app.route('/PdfIntoword', methods=['POST', 'GET'])
def fileIntoword():
    if request.method =='POST':
        files = request.files['pdffile']
        if files.name=='':
            flash("No selected files")
            return render_template('pdf to word preview.html')
        else:
            tmpdir = tempfile.mkdtemp()
            file_path=os.path.join(tmpdir, files.name)
            files.save(file_path)
            word_filePath= file_path + '.docx'
            parse(file_path,word_filePath,start=0,end=None)
            return send_from_directory(tmpdir, os.path.basename(word_filePath), as_attachment=True)
    
@app.route('/wordtopdf', methods=['POST'])
def wordToPdf():
        pythoncom.CoInitialize()
        if request.method == 'POST':
            f = request.files['file']
            if f and allowed_file(f.filename):
                    file = f.filename
                    tmp_dir = tempfile.mkdtemp()
                    uploaded_file_path = os.path.join(tmp_dir, file)
                    f.save(uploaded_file_path)
                    pdf_file_path = uploaded_file_path + '.pdf'
                
                    docx2pdf.convert(uploaded_file_path, pdf_file_path)
    
                    return send_from_directory(tmp_dir, os.path.basename(pdf_file_path), as_attachment=True)
            else:
                flash("Invalid file format. Please upload a .doc or .docx file.")
                return redirect('/wordtopdf_Page')
                     
@app.route('/pdf')
def pdf():
    return render_template('word to pdf preview.html')

@app.route('/split')
def split():
         return render_template("downsite.html")
    # return render_template("split preview.html")

@app.route('/merge')
def merge():
         return render_template("downsite.html")
    # return render_template("merge preview.html")

@app.route('/wordtopdf_Page')
def wordtopdf_Page():
    return render_template('word to pdf preview.html')

@app.route('/login')
def login():
         return render_template("downsite.html")
    # return render_template("login.html")

def loginwithGoogle():
    return redirect(url_for('index'))

def loginwithFacebook():
    return redirect(url_for('index'))

def emailLogin():
    return redirect(url_for('index'))

@app.route('/compress')
def compress():
         return render_template("downsite.html")
    # return render_template("compress preview.html")

# pdf to word
@app.route('/PdftoWord')
def PdftoWord():
    return render_template("pdf to word preview.html")

# pdf to powerpoint
@app.route('/Pdftopower')
def Pdftopower():
         return render_template("downsite.html")
    # return render_template("pdf to powerpoint preview.html")

# pdf to excel
@app.route('/pdfToExcel')
def pdfToExcel():
         return render_template("downsite.html")
    # return render_template("pdf to excel preview.html")

# word to pdf
@app.route('/wordtopdf')
def wordtopdf():
    return render_template("word to pdf preview.html")

# powerpoint to pdf
@app.route('/powerpointtopdf')
def powerpointtopdf():
         return render_template("downsite.html")
    # return render_template("powerpoint to pdf preview.html")

# excel to pdf
@app.route('/Exceltopdf')
def Exceltopdf():
         return render_template("downsite.html")
    # return render_template("excel to pdf preview.html")

@app.route('/editPdf')
def editPdf():
         return render_template("downsite.html")
    # return render_template("edit pdf options preview.html")

# pdf to jpg
@app.route('/pdftoJpg')
def pdftoJpg():
         return render_template("downsite.html")
    # return render_template("pdf to jpg preview.html")

# jpg to pdf
@app.route('/JPGtopdf')
def JPGtopdf():     
    return render_template("JPG to pdf preview.html")

@app.route('/sign')
def sign():
         return render_template("downsite.html")
    # return render_template("sign pdf preview.html")

@app.route('/watermark')
def watermark():
         return render_template("downsite.html")
    # return render_template("watermark preview.html")

@app.route('/rotate')
def rotate():
        return render_template("downsite.html")
    # return render_template("rotate pdf preview.html")
# html to pdf
@app.route('/htmltoPdf')
def htmltoPdf():
         return render_template("downsite.html")
    # return render_template("html to pdf preview.html")

@app.route('/unlockPDf')
def unlock_pDf():
        # return render_template("downsite.html")
    return render_template("unlockpdf preview.html")

@app.route('/protectPdf')
def protectPdf():
        # return render_template("downsite.html")
    return render_template("protectpdf preview.html")

@app.route('/organize')
def organize():
        return render_template("downsite.html")
    # return render_template("organize pdf preview.html")

def pdfa():
        return render_template("downsite.html")
    # return render_template("pdf to pdfa preview.html")

@app.route('/repair')
def repair():
        return render_template("downsite.html")
    # return render_template("Repair preview.html")

@app.route('/pageNum')
def pageNum():
        return render_template("downsite.html")
    # return render_template("pagenumberpreview.html")

@app.route('/ocr')
def ocr():
    return render_template("downsite.html")
    return render_template("ocr pdf preview.html")

if __name__ == '__main__':
    app.run(debug=True)
