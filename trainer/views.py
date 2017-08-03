from django.views import generic
from django.views.generic.edit import CreateView,UpdateView,DeleteView
from django.core.urlresolvers import reverse_lazy
from django.http import HttpResponse
from django.shortcuts import get_object_or_404
from django.shortcuts import render
import zipfile
import glob, os
from io import StringIO
from django.http import HttpResponse
from .models import Trainer, Curriculum, Data, Document
from .forms import DocumentForm
from reportlab.pdfgen import canvas
from io import BytesIO
import datetime
from num2words import num2words
from mmap import mmap,ACCESS_READ
from xlrd import open_workbook
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen.canvas import Canvas
from fpdf import Template
from django.core.files.storage import FileSystemStorage
import time
import shutil
import os
import smtplib, email
import xlrd
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class IndexView(generic.ListView):
    template_name='trainer/index.html'
    context_object_name='all_trainers'

    def get_queryset(self):
        return Trainer.objects.all()

class DetailView(generic.DetailView):
    model=Trainer
    template_name='trainer/details.html'
    context_object_name = 'trainer'

class TrainerCreate(CreateView):
    model=Trainer
    fields=['name', 'location', 'technology','email', 'contact']

class TrainerUpdate(UpdateView):
    model=Trainer
    fields=['name', 'location', 'technology','email', 'contact']

class TrainerDelete(DeleteView):
    model=Trainer
    success_url= reverse_lazy('trainer:index')

def search(request):
    trainer=[]
    query = request.GET['q']
    if(Trainer.objects.filter(name__icontains=query)):
        trainer+=Trainer.objects.filter(name__icontains=query)
    if (Trainer.objects.filter(location__icontains=query)):
        trainer += Trainer.objects.filter(location__icontains=query)
    if (Trainer.objects.filter(technology__icontains=query)):
        trainer += Trainer.objects.filter(technology__icontains=query)
    return render(request,'trainer/index.html', {'all_trainers': trainer, 'query': query})



def download(request):
    import zipfile
    filef = zipfile.ZipFile("test.zip", "w")
    var = request.POST.getlist('checks')
    var2= request.POST.get('radio')

    pdfprofile = []
    for i in var:
        trainer= Trainer.objects.get(pk=i)
        tech = trainer.technology
        cur = Curriculum.objects.get(name=tech)
        cur = cur.course_curriculum
        pdfprofile.append(trainer.trainer_profile)
        if cur not in pdfprofile:
            if var2=='yes':
                pdfprofile.append(cur)
        print(cur)

    for i in pdfprofile:
        filef.write("D:\\Django\\myproduct\\media\\"+str(i))
    filef.close()
    trainer = Trainer.objects.all()

    response = HttpResponse(open("test.zip", "rb").read(), content_type='application/zip')
    response['Content-Disposition'] = 'attachment; filename=test.zip'
    return response



def proposalform(request):
    return render(request, 'trainer/proposal.html')

def proposalform2(request):
    client = request.POST.get('clientname')
    course = request.POST.get('coursename')
    coursetype = request.POST.get('select')
    return render(request, 'trainer/proposal2.html', {'client': client , 'course': course, 'coursetype': coursetype})


def proposal(request):
    # Create the HttpResponse object with the appropriate PDF headers.
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="Proposal.pdf"'
    client = request.POST.get('clientname')
    course= request.POST.get('coursename')
    coursetype = request.POST.get('select')
    rate= request.POST.get('rate')
    qty = request.POST.get('qty')
    tnc1 = request.POST.get('tnc1')
    tnc2 = request.POST.get('tnc2')
    tnc3 = request.POST.get('tnc3')
    buffer = BytesIO()

    date = datetime.date.today()
    date=str(date)
    p = canvas.Canvas(buffer)
    p.setLineWidth(.3)
    p.setFont('Helvetica-Bold', 15)

    p.line(50, 720, 550, 720)
    p.line(50, 720, 50, 50)
    p.line(50, 50, 550, 50)
    p.line(550, 50, 550, 720)
    p.line(50, 720, 550, 720)
    p.line(50, 650, 550, 650)
    p.line(200, 720, 200, 470)
    p.line(50, 620, 550, 620)
    p.line(50, 500, 550, 500)
    p.line(50, 470, 550, 470)
    p.line(50, 440, 550, 440)


    p.setFont('Helvetica', 10)
    p.drawString(70, 700, "To: ")
    p.drawString(70, 685, client)
    p.setFont('Helvetica', 12)
    p.drawString(250, 690, "Quotation No: MS/2016-17/001 ")
    p.drawString(250, 670, "Date: "+date)
    p.setFont('Helvetica-Bold', 15)
    p.drawString(70, 630, "Particulars")


    if (coursetype in "Online"):
        p.setFont('Helvetica-Bold', 15)
        p.drawString(200, 750, 'Online Training Quotation')
        p.line(195, 745, 400, 745)

        p.line(280, 650, 280, 500)
        p.line(400, 650, 400, 500)
        p.line(460, 650, 460, 500)

        p.setFont('Helvetica-Bold', 12)
        p.drawString(235, 632, "Unit")
        p.drawString(290, 632, "Rate per hour(INR)")
        p.drawString(410, 632, "Hours")
        p.drawString(470, 632, "Total(INR)")

        p.setFont('Helvetica', 12)
        p.drawString(70, 600, course + " Training")
        p.drawString(70, 585, "(As per course ")
        p.drawString(70, 570, "curriculum provided) ")
        p.drawString(70, 555, "Delivery: ")
        p.drawString(70, 540, "Online Instructor Led")

        p.drawString(235, 575, "Hour")

        p.drawString(320, 575, rate+"/-")

        p.drawString(410, 575, qty+" Hrs.")
        rate1=0
        qty1=0
        for i in range(len(rate)):
            rate1 = rate1 * 10 + (ord(rate[i]) - ord('0'))
        for i in range(len(qty)):
            qty1 = qty1 * 10 + (ord(qty[i]) - ord('0'))
        total=rate1*qty1
        wtotal=num2words(total)
        wtotal=wtotal.capitalize()
        tot=str(total)
        p.drawString(470, 575, "Total: "+ tot+"/-")

        p.setFont('Helvetica-Bold', 15)
        p.drawString(70, 480, "TOTAL: ")
        p.drawString(260, 480, tot+"/-")
        p.setFont('Helvetica-Bold', 12)
        p.drawString(70, 450, "TOTAL: "+ wtotal+"/-")

        p.setFont('Helvetica-Bold', 17)
        p.drawString(70, 410, "Terms & Conditions: ")
        p.line(65, 408, 250, 408)

        p.setFont('Helvetica', 12)
        p.drawString(90, 380, " Payment: 100 % within 60 Days")
        p.drawString(90, 360, " Taxes: Service Tax as applicable")
        p.drawString(90, 340, " Training Mode: Online instructor led live ")
        p.drawString(90, 320, " Every participant must have outlook email account for cloud labs")

        if (tnc1 and tnc2 and tnc3):
            p.drawString(90, 300, " " + tnc1)
            p.drawString(90, 280, " " + tnc2)
            p.drawString(90, 260, " " + tnc3)
        elif (tnc2 and tnc3):
            p.drawString(90, 300, " " + tnc2)
            p.drawString(90, 280, " " + tnc3)
        elif (tnc1 and tnc3):
            p.drawString(90, 300, " " + tnc1)
            p.drawString(90, 280, " " + tnc3)
        elif (tnc1 and tnc2):
            p.drawString(90, 300, " " + tnc1)
            p.drawString(90, 280, " " + tnc2)
        elif (tnc1):
            p.drawString(90, 300, " " + tnc1)
        elif (tnc2):
            p.drawString(90, 300, " " + tnc2)
        elif (tnc3):
            p.drawString(90, 300, " " + tnc3)


    elif(coursetype in "Classroom"):
        p.setFont('Helvetica-Bold', 15)
        p.drawString(200, 750, 'Classroom Training Quotation')
        p.line(195, 745, 420, 745)

        days = request.POST.get('days')
        p.line(260, 650, 260, 500)
        p.line(340, 650, 340, 500)
        p.line(390, 650, 390, 500)
        p.line(445, 650, 445, 500)
        p.line(50, 440, 550, 440)

        # to be changed accordingly
        p.setFont('Helvetica-Bold', 12)
        p.drawString(220, 632, "Unit")
        p.drawString(270, 632, "No. of Days")
        p.drawString(350, 632, "Rate")
        p.drawString(395, 632, "Hrs/day")
        p.drawString(450, 632, "Course Charges")

        p.setFont('Helvetica', 12)
        p.drawString(70, 600, course+" Training")
        p.drawString(70, 585, "(As per course ")
        p.drawString(70, 570, "curriculum provided) : ")

        p.drawString(210, 590, "Per")
        p.drawString(210, 575, "Training")
        p.drawString(210, 560, "Day ")

        p.drawString(300, 570, days)
        p.drawString(290, 555, "days")

        p.drawString(342, 580, rate+"/-")
        p.drawString(342, 565, "per day")

        p.drawString(395, 580, qty+" hrs")
        p.drawString(395, 565, "per day")

        rate1 = 0
        days1 = 0
        for i in range(len(rate)):
            rate1 = rate1 * 10 + (ord(rate[i]) - ord('0'))
        for i in range(len(days)):
            days1 = days1 * 10 + (ord(days[i]) - ord('0'))
        total = rate1 * days1
        tot = str(total)

        p.setFont('Helvetica-Bold', 12)
        wtotal=num2words(total)
        wtotal=wtotal.capitalize()
        p.drawString(460, 575, tot+"/-")

        p.setFont('Helvetica-Bold', 15)
        p.drawString(70, 480, "TOTAL: ")
        p.drawString(260, 480, tot+"/-")
        p.setFont('Helvetica', 12)
        p.drawString(70, 450, "TOTAL: " + wtotal+"/-")

        p.setFont('Helvetica-Bold', 17)
        p.drawString(70, 410, "Terms & Conditions: ")
        p.line(65, 408, 250, 408)

        p.setFont('Helvetica', 12)
        p.drawString(90, 380, " Payment: 100 % within 30 Days")
        p.drawString(90, 360, " Taxes: Service Tax 15 % extra")
        p.drawString(90, 340, " Training Mode(Classroom): Classroom training for 4 days at ")
        p.drawString(100, 320, "Microsemi premises in Bangalore")
        p.drawString(90, 300, " Venue and Infrastructure for classroom training to be provided by")
        p.drawString(100, 280, " Microsemi in Bangalore")

        if(tnc1 and tnc2 and tnc3):
            p.drawString(90, 260, " "+tnc1)
            p.drawString(90, 240, " "+tnc2)
            p.drawString(90, 220, " "+tnc3)
        elif(tnc2 and tnc3):
            p.drawString(90, 260, " " + tnc2)
            p.drawString(90, 240, " " + tnc3)
        elif(tnc1 and tnc3):
            p.drawString(90, 260, " " + tnc1)
            p.drawString(90, 240, " " + tnc3)
        elif (tnc1 and tnc2):
            p.drawString(90, 260, " " + tnc1)
            p.drawString(90, 240, " " + tnc2)
        elif(tnc1):
            p.drawString(90, 260, " " + tnc1)
        elif (tnc2):
            p.drawString(90, 260, " " + tnc2)
        elif (tnc3):
            p.drawString(90, 260, " " + tnc3)

    # Draw things on the PDF. Here's where the PDF generation happens.
    # See the ReportLab documentation for the full list of functionality.
    suma = (7 * 75675678567856785) * 70
    suma = str(suma)
    resta = 100 - 9
    resta = str(resta)

    #p.drawString(60, 600, suma)
    #p.drawString(60, 590, resta)

    # Close the PDF object cleanly.
    p.showPage()
    p.save()

    # Get the value of the BytesIO buffer and write it to the response.
    pdf = buffer.getvalue()
    buffer.close()
    response.write(pdf)
    return response

def certificate(request):
    return render(request, 'trainer/certificate.html')


#Generating Single Certificate

def certificategenerate(request):
    # Create the HttpResponse object with the appropriate PDF headers.
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="Certificate.pdf"'
    name = request.POST.get('name')
    course = request.POST.get('course')
    certiid = request.POST.get('certiid')
    l=len(name)
    l=l/2
    l2 = len(course)
    l2 = l2 / 2
    l=l*2.5
    l2=l2*2.5
    date = datetime.date.today()
    date = str(date)
    buffer = BytesIO()
    logo = ImageReader('../myproduct/media/logo.png')
    date = datetime.date.today()
    date=str(date)
    p = canvas.Canvas(buffer)
    #p.setLineWidth(.3)
    p.setFont('Helvetica', 20)
    p.drawImage(logo, 0, 0,width=600,height=400)
    p.drawString(280-l,195,name)
    p.drawString(280-l2, 130, course)
    p.setPageSize((600, 400))

    p.setFont('Helvetica', 7)
    p.drawString(69, 61, date)
    p.drawString(102, 49, certiid)
    p.showPage()
    p.save()

    # Get the value of the BytesIO buffer and write it to the response.
    pdf = buffer.getvalue()
    buffer.close()
    response.write(pdf)
    return response

#Generating Multiple Pdf Certificates from Excel

def uploadcertificate(request):
    d= Document.objects.get()
    book = open_workbook('../myproduct/media/'+str(d.document))
    sheet = book.sheet_by_index(0)
    num_row = sheet.nrows - 1
    num_cell = sheet.ncols - 1
    time_folder = time.strftime("%Y-%m-%d")
    for rw in range(1, sheet.nrows):
        name, email, grade, course, lic_no = [data.value for data in sheet.row(rw)]
        print (name + " " + email + " " + course + " " + lic_no)
        # print lic_no
        if not os.path.exists('D:/Certificate_App/finalcert-master/Backup_Certificate/%s' % (time_folder,)):
            os.makedirs('D:/Certificate_App/finalcert-master/Backup_Certificate/%s' % (time_folder,))
        if not os.path.exists('D:/Certificate_App/finalcert-master/Backup_Certificate/%s/%s' % (time_folder, course,)):
            os.makedirs('D:/Certificate_App/finalcert-master/Backup_Certificate/%s/%s' % (time_folder, course,))
            # this will define the ELEMENTS that will compose the template.
        elements = [
            {'name': 'company_logo', 'type': 'I', 'x1': 68.0, 'y1': 9.0, 'x2': 150.0, 'y2': 38.0, 'font': None,
             'size': 0.0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I',
             'text': 'logo', 'priority': 2, },
            {'name': 'bg1', 'type': 'I', 'x1': 44.0, 'y1': 38.0, 'x2': 180.0, 'y2': 50.0, 'font': None, 'size': 0.0,
             'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': 'logo',
             'priority': 2, },
            {'name': 'title', 'type': 'T', 'x1': 63.0, 'y1': 42, 'x2': 168.0, 'y2': 47, 'font': 'Arial', 'size': 18.0,
             'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0XFFFFFF, 'background': 0, 'align': 'I', 'text': '',
             'priority': 3, },
            {'name': 'to', 'type': 'T', 'x1': 70.0, 'y1': 58, 'x2': 150.0, 'y2': 62, 'font': 'Times', 'size': 24.0,
             'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '',
             'priority': 2, },
            # { 'name': 'ul1', 'type': 'T', 'x1': 69.5, 'y1': 61.7, 'x2': 200.0, 'y2': 61.7, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
            {'name': 'complete', 'type': 'T', 'x1': 60.0, 'y1': 68, 'x2': 168.0, 'y2': 105.0, 'font': 'Times',
             'size': 24.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'C',
             'text': '', 'priority': 2, },
            # { 'name': 'ul2', 'type': 'T', 'x1': 80, 'y1': 73, 'x2': 200.0, 'y2': 73, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
            # { 'name': 'grade', 'type': 'T', 'x1': 78.0, 'y1': 85, 'x2': 130, 'y2': 85, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
            #	{ 'name': 'date', 'type': 'T', 'x1': 70.0, 'y1': 99, 'x2': 140, 'y2': 99, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
            {'name': 'bg', 'type': 'I', 'x1': 0, 'y1': 100, 'x2': 216.0, 'y2': 154, 'font': None, 'size': 0.0,
             'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': 'logo',
             'priority': 2, },
            {'name': 'sign', 'type': 'T', 'x1': 145, 'y1': 137, 'x2': 197, 'y2': 142, 'font': 'Times', 'size': 10.0,
             'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '',
             'priority': 3, },
            {'name': 'name', 'type': 'T', 'x1': 75, 'y1': 68, 'x2': 150, 'y2': 78, 'font': 'Times', 'size': 22.0,
             'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 1, 'align': 'C', 'text': '',
             'priority': 3, },
            {'name': 'course', 'type': 'T', 'x1': 79, 'y1': 95, 'x2': 150, 'y2': 110, 'font': 'Times', 'size': 22.0,
             'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'C', 'text': '',
             'priority': 3, },
            # { 'name': 'grade1', 'type': 'T', 'x1': 95, 'y1': 80, 'x2': 150, 'y2': 86, 'font': 'Times', 'size': 18.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
            {'name': 'on', 'type': 'T', 'x1': 15, 'y1': 150, 'x2': 150, 'y2': 100, 'font': 'Arial', 'size': 8.0,
             'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '',
             'priority': 3, },
            {'name': 'copywrite', 'type': 'T', 'x1': 74, 'y1': 140, 'x2': 150, 'y2': 145, 'font': 'Times', 'size': 8.0,
             'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '',
             'priority': 3, },
            {'name': 'license', 'type': 'T', 'x1': 15, 'y1': 114, 'x2': 150, 'y2': 145, 'font': 'Arial', 'size': 8.0,
             'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '',
             'priority': 3, },
        ]

        # here we instantiate the template and define the HEADER
        f = Template(format="A5", orientation="L", elements=elements,
                     title="Certificate")
        f.add_page()

        # we FILL some of the fields of the template with the information we want
        # note we access the elements treating the template instance as a "dict"
        f["company_logo"] = "D:/Certificate_App/finalcert-master/index.png"
        f["bg1"] = "D:/Certificate_App/finalcert-master/bg1.jpg"
        f["bg"] = "D:/Certificate_App/finalcert-master/bg2.png"
        f["title"] = "CERTIFICATE OF COMPLETION"
        f["to"] = "This is to certify that"
        f["sign"] = "Sanjay Verma, Founder & CEO"
        f["ul1"] = "____________________________________________________"
        f["ul2"] = "_______________________________________________"
        f["complete"] = "has satisfactorily completed a course on"
        f["copywrite"] = "@2017 Blue Camphor Technologies (P) Ltd"
        f["grade"] = "with  ________   Grade"
        f["date"] = "Awarded on  ______________"
        f["name"] = name
        f["course"] = course
        f["grade1"] = grade
        f["on"] = "Date - " + time.strftime("%d/%m/%Y")
        f["license"] = "Certificate ID - " + lic_no

        # and now we render the page
        f.render("D:/Certificate_App/finalcert-master/Backup_Certificate/%s/%s/%s.pdf" % (time_folder, course, name,))

        course_url = course.replace(" ", "%20")
        # print course_url
        lic_date = time.strftime("%Y%m")
    # print lic_date


        # email code
        fromaddr = "support@skillspeed.com"
        print(name + " " + email + " " + course + " " + lic_no)
        toaddr = email
        name_array = name.split()
        # print name_array[0]

        msg = MIMEMultipart('alternative')
        #cc_1 = "shnehil@skillspeed.com"
        #cc_2 = "gajal@skillspeed.com"
        msg['From'] = fromaddr
        msg['To'] = toaddr
        #msg['cc'] = cc_1 + "," + cc_2
        msg['Subject'] = "[CERTIFICATE]: " + course

        body = """\
    	<html>
    		<head></head>
    		<body>
    		<p>
            Dear  %s,
            <br></br>
            <br></br>
            We thank you for attending our <b> %s </b> & hope you enjoyed the sessions much as we enjoyed teaching you.
            <br></br>
            <br></br>
            This mail certifies that <b> %s </b> has successfully completed the  <b> %s</b>.
            <br></br>
            <br></br>
            Add your certificate to Linkedin
            <div><a href="https://www.linkedin.com/profile/add?_ed=0_JhwrBa9BO0xNXajaEZH4q9ZriGQBiq56O8XQeptEb_xAD6iVbtTHBphjlBeRBwz4aSgvthvZk7wTBMS3S-m0L6A6mLjErM6PJiwMkk6nYZylU7__75hCVwJdOTZCAkdv&pfCertificationName=%s&pfLicenseNo=%s&pfCertStartDate=%s&trk=onsite_html" rel="nofollow" target="_blank"><img src="https://download.linkedin.com/desktop/add2profile/buttons/en_US.png" alt="LinkedIn Add to Profile button"></a></div>
            <br></br>
            <br></br>
            If you have any queries, feel free to reach out to our support team at support@skillspeed.com
            <br></br>
            USA & Rest of World: +1-661-241-4796
            <br></br>
            India: +91-906-602-0904
            <br></br>
            <br></br>
            <b>Support</b>
            <br></br>
            <br></br>
            <b>Good Karma</b>
            <br></br>
            <br></br>
            Please do spread the word about our sessions and like us on Facebook & leave a comment with your thoughts.
            <br></br>
            <div><a href="http://www.facebook.com/SkillspeedOnline"><img src="http://www.mail-signatures.com/articles/wp-content/themes/emailsignatures/images/facebook-35x35.gif"></a><a href="http://www.linkedin.com/company/skillspeed"><img src="http://www.mail-signatures.com/articles/wp-content/uploads/2014/08/linkedin.png" width="35" height="35"></a></div>
            <br></br>
            <b>Thanks & Regards,
            <br></br>
            Skillspeed Support Team
            <br></br></b>
            </p>
            </body>
            </html>
            """ % (name, course, name, course, course_url, lic_no, lic_date)

        msg.attach(MIMEText(body, 'html'))

        filename = "certificate.pdf"
        attachment = open(
            "D:/Certificate_App/finalcert-master/Backup_Certificate/%s/%s/%s.pdf" % (time_folder, course, name,), "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)
        #print(msg)
        toadd = [toaddr]
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(fromaddr, "skillspeed")
        text = msg.as_string()
        Document.objects.all().delete()
        server.sendmail(fromaddr, toadd, text)
        server.quit()

        #shutil.rmtree('../myproduct/media/sheets')

    return render(request, 'trainer/certificate.html')

#Upload Excel file to Database

def excelupload(request):
    Document.objects.all().delete()

    if request.method == 'POST':
        form = DocumentForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
    else:
        form = DocumentForm()
        print(form)
    return render(request, 'trainer/certificate.html', {
        'form': form
    })


#UPLOADING bulk data to Data database

def upload(request):

    book = open_workbook('../myproduct/Datasheet.xlsx')
    sheet = book.sheet_by_index(0)
    col=sheet.ncols
    print(sheet.nrows)
    for row_index in range(sheet.nrows):
        col = sheet.ncols
        d=Data()
        d.name=(sheet.cell(row_index, sheet.ncols-col).value)
        col-=1
        d.contact = (sheet.cell(row_index, sheet.ncols-col).value)
        col-=1
        d.email=(sheet.cell(row_index, sheet.ncols-col).value)
        col-=1
        d.technology = (sheet.cell(row_index, sheet.ncols-col).value)
        exist=Data.objects.filter(name=d.name)
        print(row_index)
        print(d.name)
        if(not exist):
            d.save()

    return render(request, 'trainer/index.html', {'all_trainers': Data.objects.all()})
