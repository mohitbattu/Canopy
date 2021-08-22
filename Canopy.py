import tabula
import PyPDF2
from pdf2image import convert_from_path
import cv2
import pytesseract
from pdfminer.converter import TextConverter
from PIL import Image
from pytesseract import Output
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTFigure, LTTextBox
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFParser
from io import StringIO
import xlsxwriter
filepath=r"C:/Users/raobk/Desktop/Python_Technical_Assignment/test_input.pdf"
images = convert_from_path("C:/Users/raobk/Desktop/Python_Technical_Assignment/test_input.pdf",poppler_path=r'C:/poppler-0.68.0/bin')
for i, image in enumerate(images):
    fname = 'image'+str(i)+'.png'
    image.save(fname, "PNG")
img = cv2.imread('image0.png')
gray_img=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
threshold_img=cv2.threshold(gray_img,0,255,cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
custom_config = r'--oem 3 --psm 6'
d = pytesseract.image_to_data(threshold_img, output_type=Output.DICT, config = custom_config,lang='eng')
n_boxes = len(d['text'])
for i in range(n_boxes):
    try:
        if float(d['conf'][i]) > 60.0:
            (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
            threshold_img = cv2.rectangle(threshold_img, (x, y), (x + w, y + h), (0, 255, 0), 2)
            # print(x,y)
    except ValueError as e:
        print(e)
# cv2.imshow('img', threshold_img)
# cv2.waitKey(0)
print(d['text'])
parse_text = []
word_list = []
last_word= ""
for word in d['text']:
    if word!="":
        word_list.append(word)
        last_word=word
    if(last_word!='' and word== '') or (word==d['text'][-1]):
        parse_text.append(word_list)
        word_list=[]


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text

stored=convert_pdf_to_txt(filepath)
gettingdata=stored.split("\n")
def check_if_empty(data):
    if data!= None and data!='':
        return False
    elif data=='' or data==" ":
        return True
# print(gettingdata)
added=[]
count=0
for i in range(count,len(gettingdata)):
    try:
        if check_if_empty(gettingdata[i])==False and check_if_empty(gettingdata[i+1])==False:
            if gettingdata[i]!='Deposits':
                added.append(gettingdata[i]+gettingdata[i+1])
                if(len(gettingdata)==count):
                    count+=2
        elif (check_if_empty(gettingdata[i])==False and check_if_empty(gettingdata[i-1])==True and check_if_empty(gettingdata[i+1])==True):
            added.append(gettingdata[i])
        if gettingdata[i]=='Deposits' and gettingdata[i+1]=="Deposits AUD":
            added.append(gettingdata[i])
            added.append(gettingdata[i+1])
    except:
        print("Error")
# print(added)
modifieddata=[]
for i in range(len(added)):
    if added[i]!=" ":
        modifieddata.append(added[i])

def is_integer_num(n):
    if isinstance(n, int):
        return True
    if isinstance(n, float):
        return n.is_integer()
    return False
# check(parse_text,storedata)
print(modifieddata)
new=[]
copy_text=parse_text
count=0
p=0
for i in range(len(copy_text)):
    if i==29:
        break
    for j in range(len(copy_text[i])):
        if len(copy_text[i])>1:
            if (int(len(copy_text[i])-1))==j or j>=int(len(copy_text[i])):
                break
            if(j>0):
                result4=copy_text[i][j-1][0].isupper()
            result1=copy_text[i][j][0].isupper()
            
            result2=copy_text[i][j+1][0].isupper()
            
            result3=copy_text[i][j+1][0+1].isupper()
            
            try:
                result7=copy_text[i][j+2][0].isupper()
            except:
                print('Error')
            try:
                if(count==2 and copy_text[i][j+2][0].isupper()==True and len(copy_text[i])==16):
                    for k in range(len(modifieddata)):
                        if copy_text[i][j]==modifieddata[k][:len(copy_text[i][j])]:
                            parse_text[i][j]=modifieddata[k]
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text.remove(parse_text[i+1])
                            print(parse_text)
                            count=0
                            p=1
                            break
            except:
                print("error")
            if(count==6 and result4==False and result7==True and len(copy_text[i])==9):
                parse_text[i][j]=copy_text[i][j]+" "+copy_text[i][j+1]+" "+copy_text[i][j+2]+" "+copy_text[i][j+3]+" "+copy_text[i][j+4]+" "+copy_text[i][j+5]
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                print(parse_text)
                count=0
                break
            if(count==4 and result7==True and result4==False and len(copy_text[i])==9):
                parse_text[i][j]=copy_text[i][j]+" "+copy_text[i][j+1]+" "+copy_text[i][j+2]+" "+copy_text[i][j+3]+" "+copy_text[i][j+4]+" "+copy_text[i][j+5]
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                parse_text[i].remove(parse_text[i][j+1])
                print(parse_text)
                count=0
                break
            try:
                if(count==4 and copy_text[i][j+2][0].isupper()==True and len(copy_text[i])==15):
                    if(result7==True and result4==False and copy_text[i][j+2]=='GOLDMAN'):
                        copy_text[i][j]=copy_text[i][j]+" "+copy_text[i][j+1]+" "+copy_text[i][j+2]+" "+copy_text[i][j+3]+" "+copy_text[i][j+4]+" "+copy_text[i][j+5]+" "+copy_text[i][j+6]+" "+copy_text[i][j+7]+" "+copy_text[i][j+8]+" "+copy_text[i][j+9]+" "+copy_text[i][j+10]+" "+copy_text[i][j+11]+" "+copy_text[i+1][0]+" "+copy_text[i+1][1]+" "+copy_text[i+1][2]
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text.remove(parse_text[i+1])
                        print(parse_text)
                        count=0
                        break
                    for k in range(len(modifieddata)):
                        if copy_text[i][j]==modifieddata[k][:len(copy_text[i][j])]:
                            parse_text[i][j]=modifieddata[k]
                            print(parse_text)
                            print(j)
                            print(copy_text[i])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            
                            parse_text.remove(parse_text[i+1])
                            print(parse_text)
                            count=0
                            p=1
                            break
            except:
                print('Error')
            
            try:
                if(count==2 and result7==True and result4==False and len(copy_text[i])==15 and copy_text[i][j+3]=='HKD'):
                    parse_text[i][j]=copy_text[i][j]+" "+copy_text[i][j+1]+" "+copy_text[i][j+2]+" "+copy_text[i][j+3]+" "+copy_text[i][j+4]+" "+copy_text[i][j+5]+" "+copy_text[i][j+6]+" "+copy_text[i][j+7]+" "+copy_text[i][j+8]+" "+copy_text[i][j+9]+" "+copy_text[i][j+10]+" "+copy_text[i][j+11]+" "+copy_text[i+1][0]+" "+copy_text[i+1][1]+" "+copy_text[i+1][2]+" "+copy_text[i+1][3]+" "+copy_text[i+1][4]+" "+copy_text[i+1][5]
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text.remove(parse_text[i+1])
                    print(parse_text)
                    count=0
                  
                    break

                    
            except:
                print('Error')
            
            try:
                if(count==4 and len(copy_text[i])==18 and result4==False):
                    for k in range(len(modifieddata)):
                        if copy_text[i][j]==modifieddata[k][:len(copy_text[i][j])]:
                            parse_text[i][j]=modifieddata[k]
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            count=0
                            p=1
                            break
            except:
                print('Error')

            try:
                if(count==2 and len(copy_text[i])==15 and result4==False and copy_text[i][j]=='CONSENT'):
                    for k in range(len(modifieddata)):
                        if copy_text[i][j]==modifieddata[k][:len(copy_text[i][j])]:
                            parse_text[i][j]=modifieddata[k]
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            parse_text[i].remove(parse_text[i][j+1])
                            count=0
                            p=1
                            break
            except:
                print('Error')
                          
                    
            try:
                result5=is_integer_num(int(copy_text[i][j]))
                result6=is_integer_num(int(copy_text[i][j+2]))
                if(result5==True and result2==True and result6==True):
                    parse_text[i][j]=copy_text[i][j]+ " " +copy_text[i][j+1] + " "+copy_text[i][j+2]
                    parse_text[i].remove(parse_text[i][j+1])
                    parse_text[i].remove(parse_text[i][j+1])
                    count+=1
                    print(count)
                    print(parse_text)
            except:
                print('Error')
            


            if(result1==True and result2==False and result3==False):

                parse_text[i][j]=copy_text[i][j]+" "+copy_text[i][j+1]
                parse_text[i].remove(parse_text[i][j+1])
                print(parse_text)
                
            elif(j >0 and result4==True and result1==False):
                parse_text[i][j]=copy_text[i][j-1]+" "+copy_text[i][j]
                parse_text[i].remove(parse_text[i][j-1])
                        # copy_text[i].remove(copy_text[i][j+1])
                print(parse_text)
            elif(result1==True and result2==True and result3==True):
                if p==0:
                    if result4==False and result7==True:
                        parse_text[i][j]=copy_text[i][j]+" "+copy_text[i][j+1]+ " "+ copy_text[i][j+2] + " "+ copy_text[i][j+3]
                        if(count==4 and len(copy_text[i])==8 and result4==False):
                            parse_text[i][j+4]=copy_text[i][j+4]+copy_text[i][j+5]
                            parse_text[i].remove(parse_text[i][j+5])
                        if(count==2 and len(copy_text[i])==8 and result4==False):
                            parse_text[i][j+4]=copy_text[i][j+4]+copy_text[i][j+5]
                            parse_text[i].remove(parse_text[i][j+5])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        parse_text[i].remove(parse_text[i][j+1])
                        print(parse_text)
                    else:
                        parse_text[i][j]=copy_text[i][j]+" "+copy_text[i][j+1]
                        parse_text[i].remove(parse_text[i][j+1])
                    # copy_text[i].remove(copy_text[i][j+1])
                        print(parse_text)
                else:
                    p=0
                    break
            
print(parse_text)
raised=0
currency=''
for i in range(len(parse_text)):
    if i==25:
        break
    for j in range(len(parse_text[i])):
        if(parse_text[i][j]=="Opening balance"):
            if(parse_text[i+1][j-1][:9]=="Month end"):
                parse_text.remove(parse_text[i])
                parse_text.remove(parse_text[i])
        print(parse_text[i][j])
        if(parse_text[i][j][3:6]=='Mar'):
            print(parse_text[i][j])
            parse_text[i][j]=parse_text[i][j][7:]+"-"+"03"+"-"+parse_text[i][j][:2]
            print(parse_text[i][j])
        elif(parse_text[i][j][3:6]=='Feb'):
            print(parse_text[i][j])
            parse_text[i][j]=parse_text[i][j][7:]+"-"+"02"+"-"+parse_text[i][j][:2]
            print(parse_text[i][j])
print(parse_text)

for i in range(len(parse_text)):
    if(i==20):
        break
    for j in range(len(parse_text[i])):
        print(parse_text[i][j])
        
        if(parse_text[i][j]=='Transaction details'):
            parse_text.remove(parse_text[i])
        if(parse_text[i][j]=='Deposits'):
            parse_text.remove(parse_text[i])
        if(parse_text[i][j]=='Deposits HKD'):
            raised=1
            currency=parse_text[i][j][9:]
            parse_text.remove(parse_text[i])
            print(parse_text)
        if(parse_text[i][j-1]=='Deposits USD'):
            raised=1
            currency=parse_text[i][j-1][9:]
            print(parse_text)
            parse_text.remove(parse_text[i])
        if(raised==1 and currency=="USD"):
            parse_text[i].insert(len(parse_text[i])-1,'USD')
            print(parse_text)
            break
        if(raised==1 and currency=='HKD'):
            if(parse_text[i][j][:9]=="Month end"):
                raised=0
                currency=''
                parse_text.remove(parse_text[i])
                continue
            parse_text[i].insert(len(parse_text[i])-1,'HKD')
            print(parse_text)
            break

def keysearch(keywords):
    list1=[]
    for i in range(len(parse_text)):
        for j in range(len(parse_text[i])):
            if parse_text[i][j]==keywords:
                list1.append(i) 
    return list1

jump=0
for i in range(len(parse_text)):
    if(jump==1):
        energy=energy-1
        print(energy)
        if(energy==0):
            jump=0
            continue
    for j in range(len(parse_text[i])):
        print(parse_text[i])
        keyword=parse_text[i][0]
        # print(keyword)
        getval=keysearch(keyword)
        updating=list(set(getval))
        if(jump==0):
            energy=len(updating)-1
        if(len(updating)>1 and jump==0):
            maximum=max(updating)
            minimum=min(updating)
            diff=maximum-minimum
            if(diff>2 and jump==0):
                value1=parse_text[maximum][len(parse_text[maximum])-1]
                value2=parse_text[minimum][len(parse_text[minimum])-1]
                try:
                    if(float(value1)>float(value2)):
                        parse_text.insert(minimum+1,parse_text[maximum])
                        parse_text.pop(maximum+1)
                        print(parse_text)
                        jump=1
                except:
                    if(value1[0]=='(' and value1[-1]==')' and is_integer_num(int(value1[1]))==True):
                        value1=value1.replace(",","")
                        value2=value2.replace(",","")
                        if(float(value1[1:-1])<float(value2)):
                            parse_text.insert(minimum,parse_text[maximum])
                            parse_text.pop(maximum+1)
                            print(parse_text)
                            jump=1
                    elif(value2[0]=='(' and value2[-1]==')' and is_integer_num(int(value2[1]))==True):
                        modif=value2[1:-1]
                        modif=modif.replace(",","")
                        print(modif)
                        value1=value1.replace(",","")
                        print(value1)
                        if(float(modif)>float(value1)):
                            parse_text.insert(minimum,parse_text[maximum])
                            parse_text.pop(maximum+1)
                            print(parse_text)
                            jump=1
                        
                        print(parse_text[i])


        break

for i in range(len(parse_text)):
    for j in range(len(parse_text[i])):
        if(parse_text[i][j]=='Amount'):
            parse_text[i][j]='Currency'
            parse_text[i].insert(j+1,'Debit')
            parse_text[i].insert(j+2,'Credit')
        data=parse_text[i][len(parse_text[i])-1]
        try:
            if(is_integer_num(int(data[0]))==True and is_integer_num(int(data[-1]))==True):
                data=data.replace(",","")
                parse_text[i][len(parse_text[i])-1]=float(data)
                print(is_integer_num(parse_text[i][j]))
                parse_text[i].insert((len(parse_text[i]))-1,'')
                break
        except:
            if(data[0]=='(' and data[-1]==')' and is_integer_num(int(data[1]))==True):
                d1=parse_text[i][len(parse_text[i])-1][1:-1]
                data=d1.replace(",","")
                parse_text[i][len(parse_text[i])-1]=float(data)
                parse_text[i].insert((len(parse_text[i])),'')
                break
            print('Error')
print(parse_text)            
book = xlsxwriter.Workbook('Canopy.xlsx')
sheet = book.add_worksheet()
for i in range(len(parse_text)):
    for j in range(len(parse_text[i])):
        sheet.write(i,j,parse_text[i][j])
book.close()
cv2.imshow('img', threshold_img)
cv2.waitKey(0)