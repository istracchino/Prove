from asyncio.windows_events import NULL
from dataclasses import replace
from logging import exception
import xml.etree.ElementTree as Xet
import os
import zipfile
from PIL import Image
import hashlib
import shutil
import os
import comtypes.client
import sys
import fnmatch
#TODO


dirProdotti = 'Ricette'
dirRicette = 'WV/Ricette/'
dirImmagini = 'WV/image/'
os.chdir(os.path.dirname(os.path.realpath(__file__)))


document = []
outlist = []
relation = {}
xmlPAR = {}











#def main():
    #df,dc,db=selProd(dirProdotti)
    #print('csvTOword',dc)
    #csvTOword(dirProdotti)
    #csvTOword(df,dc,db)
    #wordTOcsv("Sample.docx","proce.csv")
    #x,y = selProd(dirRicette)
    #for op in os.scandir(x):
        #print(op,x,y)
        #csvTOword(op,y)
        


def wordTOpdf(fname):
    wdFormatPDF = 17
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open('C:/Users/ipers/Desktop/ITS/CERRI/AUDIT/py/WORDS/rec1.DOCX')
    doc.SaveAs('C:/Users/ipers/Desktop/ITS/CERRI/AUDIT/py/WORDS/out_file.pdf', FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
def csvTOword(dirProdotti):
    RecipeFolder,name,imgpath=selProd(dirProdotti)
    openXML()
    create_document(name+'.DOCX')
    body = outdoc.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
    body.insert(0,TITOLO('PROCEDURA ASSEMBLAGGIO '+name))
    outlist=[]
    l = os.listdir(RecipeFolder)

    matches = fnmatch.filter(l, 'Fase?.csv')
    for i in matches:
        for x in open(os.path.join(RecipeFolder,i)).readlines():
            x.strip('\n')
            x=x.split(";")
            if x.__len__() >=2 and '.' in x[1]:
                outlist.append(Xet.tostring(STEPIMG(x[0],x[1].strip('\n'),imgpath)))
                outlist.append(Xet.tostring(xmlPAR['blank']))
            else:
                outlist.append(Xet.tostring(STEP(x[0])))
                outlist.append(Xet.tostring(xmlPAR['blank']))
    i=1
    for x in outlist:
        i+=1
        body.insert(i,Xet.fromstring(x))  
    save_document()
def openXML():
    global xmlrels
    global docu
    global outdoc
    global log
    log = open('log.txt',"w")
    with zipfile.ZipFile('word/Sample.docx') as zf:
        docs2 = zf.open("word/document.xml")
        docu = Xet.parse(docs2).getroot()
        
    with open('word/document.xml.rels') as docs:  #FILE 'BIANCO' PER RELAZIONI
        xmlrels = Xet.parse(docs).getroot()

    with open("word/document.xml") as docs3:  #FILE BIANCO PER DOCUMENTO
        outdoc = Xet.parse(docs3).getroot()
    SampleParagraph()
def wordTOcsv(docName,csvName):
    with zipfile.ZipFile(docName, "r") as archive:
        document_parser(archive)
        relation_parser(archive)
        extract_image(archive)
    makeCSV(csvName)
def save_document():
    OUTdocx.writestr('word/_rels/document.xml.rels',Xet.tostring(xmlrels))
    OUTdocx.writestr('word/document.xml',Xet.tostring(outdoc))
def SampleParagraph():
        body = docu.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
        par = body.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
        xmlPAR['TITOLO'] = par[0]
        xmlPAR['STEP'] = par[1]
        xmlPAR['blank'] = par[2]
        xmlPAR['STEPIMG'] = par[3]
def create_document(outDOCXpath):
    global OUTdocx
    shutil.copy('word/empty.docx',outDOCXpath)
    OUTdocx = zipfile.ZipFile(outDOCXpath,"a")
def create_relation(imgPath):
        global xmlrels
        numid = 0
        numimg = 0

        for n in xmlrels.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            numx = int(n.attrib['Id'].strip('rId'))
            if numx > numid:
                numid = numx
            if n.attrib['Type'] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                numy = int(n.attrib['Target'].strip('media/image.PNG'))
                if numy > numimg:
                    numimg = numy
        numid = numid+1
        numimg = numimg+1
        
        relation = Xet.fromstring('<ns0:Relationship xmlns:ns0="http://schemas.openxmlformats.org/package/2006/relationships" Id="'+'rId'+str(numid)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image'+str(numimg)+'.PNG" />')
        xmlrels.append(relation)
        OUTdocx.write(imgPath,'word/media/image'+str(numimg)+'.PNG',)
        return 'rId'+str(numid)      
def TITOLO(txt):
    out = xmlPAR['TITOLO']
    out.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r').find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t').text = txt
    return out
def STEP(txt):
    out = xmlPAR['STEP']
    out.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r').find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t').text = txt
    return out
def STEPIMG(txt,img,imgpath):
    xcns=0
    rId = create_relation(os.path.join(imgpath,img))
    out = xmlPAR['STEPIMG']
    find = out.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    for x in find:
        try:
            x.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t').text = txt
        except:
            xcns += 1
        try:
            x.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing').find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor').find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphic').find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData').find('{http://schemas.openxmlformats.org/drawingml/2006/picture}pic').find('{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill').find('{http://schemas.openxmlformats.org/drawingml/2006/main}blip').attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'] = rId
        except:
            xcns += 1

        


    return out
def makeCSV(fileName):
    myFile = open(fileName,"w")
    for line in document:
        img = ''
        if line[1] != None:
            img = relation[line[1]]['img']
        txt = ''
        if line[0] != None:
            txt = line[0]
        myFile.writelines(txt+";"+img+"\n")
def hashMD5(image):
    md5hash = hashlib.md5(Image.open(image).tobytes())
    return str(md5hash.hexdigest())
def extract_image(archivio):
    for x in document:
        if x[1] != None:
            cut = 'media'
            path = 'word/media'+str(relation[x[1]]['Path']).strip(cut)
            img = 'img/'+path
            archivio.extract(path,'img')
            hash = hashMD5(img)
            try:
                os.rename(img,'img/'+hash+'.png')
            except FileExistsError:
                os.replace(img,'img/'+hash+'.png')
            shutil.rmtree('img/word')
            relation[x[1]]['img'] = hash+'.png'
def relation_parser(archivio):
        doc = archivio.open("word/_rels/document.xml.rels")
        doc = Xet.parse(doc).getroot()
        for x in doc:
            if x.attrib['Type'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
                relation[x.attrib['Id']] = {"Path" : x.attrib['Target']}
def document_parser(archivio):
        doc = archivio.open("word/document.xml")
        doc = Xet.parse(doc).getroot()
        body = doc.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
        #for x in body:
        par = body[0].findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
        for x in par:
            txt = None
            img = None
            data = x.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
            for i in data:
                try:
                    txt = i.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t').text
                except AttributeError as x:
                    cv = x
                try:
                    img = i.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing').find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline').find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphic').find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData').find('{http://schemas.openxmlformats.org/drawingml/2006/picture}pic').find('{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill').find('{http://schemas.openxmlformats.org/drawingml/2006/main}blip').attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
                    #for x in img:
                        #print(x)
                except (AttributeError,TypeError) as cv:
                    #print(cv)
                    cv = 0
            document.append([txt,img])

def selProd(recPath):
    fam={}
    prod={}
    i=1
    for x in os.scandir(recPath):
        if x.is_dir() == True:
            fam[str(i)]=x.path
            print(i,'-',x.name)
            i+=1
    sel = str(input('Sel. famiglia: '))
    ii=1
    for x in os.scandir(os.path.join(fam[sel],dirRicette)):
        if x.is_dir() == True:
            prod[str(ii)]=[x.path,x.name]
            print(ii,'-',x.name)
            ii+=1
    sel2 = str(input('Sel. prodotto: '))  
    sel3 = int(input('Sel. n° Op: '))
    return os.path.join(prod[sel2][0]+'/'+str(sel3)+'Op'),prod[sel2][1],os.path.join(fam[sel],dirImmagini)
def listZIP():
    zf = zipfile.ZipFile("cestino/A.docx", "r")
    for x in zf.namelist():
        print(x)

#if __name__ == '__main__':
   #main()