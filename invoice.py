from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.units import inch
import pandas as pd
#loading data
a=pd.read_excel('invoice.xlsx')
#loading image
im=Image('logo.jpg',width=1.5*inch,height=inch,hAlign='LEFT')
#The invoice to be printed coressponding to the invoice number
b=input('Invoice Number :');
#Data processing
A=a[a["Invoice Number"] == b]
d=A.customer[A.customer.index[0]];
e=str(A.Date[A.Date.index[0]])
A=A.drop(['customer', 'Invoice Number','Date',], axis=1)
f=str(sum(A.Amount))
A=A.append({'Amount':'INR '+f,"Tax":'INR '+str(sum( A.Tax[0:-1])),"Quantity":"NOS "+str(sum(A.Quantity[:-1]))} , ignore_index=True)
A=A.round(2)
A=A.fillna(' ')
s=Spacer(0, 30)
elements = []
styles = getSampleStyleSheet()
#Creating PDF
doc = SimpleDocTemplate((str(b)) + '_' + str(d) +'.pdf')
#Printing data in the PDF
elements.append(im)
styles['Title'].spaceAfter=1
elements.append((Paragraph('<para align=right> COMPANY NAME<para/>', styles['Title'])))
elements.append((Paragraph('<para align=right>ADRESS LINE1<para/>',styles['Normal'])))
elements.append((Paragraph('<para align=right>ADRESS LINE2<para/>',styles['Normal'])))
elements.append(s)
elements.append(Paragraph("<u>INVOICE</u>", styles['Title']))
elements.append(s)
elements.append((Paragraph('Invoice Number :'+ b,styles['Normal'])))
elements.append((Paragraph('Customer Name :'+ str(d),styles['Normal'])))
elements.append((Paragraph('Date :'+ str(e),styles['Normal'])))
elements.append(s)
lista = [A.columns[:,].values.astype(str).tolist()] + A.values.tolist()
ts= [('LINEABOVE', (0,0), (-1,0), 2, colors.darkgray),
     ('LINEBELOW', (0,0), (-1,0), 2, colors.darkgray),
     ('LINEABOVE', (0,2), (-1,-1), 0.25, colors.black),
     ('LINEBELOW', (0,-1), (-1,-1), 2, colors.darkgray),
     ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
     ('ALIGN', (0,0), (-1,0), 'CENTER'),
     ('BACKGROUND', (-1,-1 ), (-1, -1), colors.green),
     ('BACKGROUND', (-2,-1 ), (-2, -1), colors.blueviolet),
     ('BACKGROUND', (-4,-1 ), (-4, -1), colors.darkkhaki),
     ('LINEABOVE', (0,-1), (-1,-1), 2, colors.darkgray)]
table = Table(lista, style=ts)
elements.append(table)
elements.append(Spacer(0,10))
elements.append(Paragraph("<para align=right>Total payable amount : <b> INR " + f+"<b/><para/>",styles['Normal']))
elements.append(s)
elements.append(Paragraph('For quries contact +91 9876543210',styles['Normal']))
elements.append(s)
elements.append(Paragraph('THANK YOU',styles['Title']))
#building the pdf
doc.build(elements)

