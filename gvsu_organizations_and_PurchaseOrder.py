import json
import uuid
import xlrd
import os
import xlwt
import xlsxwriter
import openpyxl
import os.path
import datetime
from datetime import datetime
import requests

class contacts():

    def __init__(self,idcontact, contactfirstName, contactlastName, contactemails,contactcategories):
        self.contactid=idcontact
        self.contactfirstName= contactfirstName
        self.contactlastName= contactlastName
        self.language= "en-us"
        self.contactnotes=""
        self.contactnumerotelefono=[]
        self.contactemails=contactemails
        self.contactcategories=contactcategories
        self.contactaddresses=[]
        self.contacturls=[]
        self.contactinactive= False


    def printcontacts(self,fileName,phone):
        contactarchivo=open(fileName+"_contacts.json", 'a')
        contacto={
            #"prefix": "",
            "id": self.contactid,
            "firstName": self.contactfirstName,
            "lastName": self.contactlastName,
            "language": self.language,
            "notes": self.contactnotes,
            "phoneNumbers": phone,
            "emails": self.contactemails,
            "addresses": self.contactaddresses,
            "urls": self.contacturls,
            "categories": self.contactcategories,
            "inactive": self.contactinactive,
           }
        json_contact = json.dumps(contacto)
        print('contacto:', json_contact)
        contactarchivo.write(json_contact+"\n")
        return contacto['id']
#end

class interfaces():

    def __init__(self,iI,nI,uI,NI,aI,tI):
        self.interid=iI
        self.intername = nI
        self.interuri = uI
        self.deliveryMethod= "Online"
        self.notes= NI
        self.interavailable=True
        self.intertype=tI
        
    def printinterfaces(self, fileName):
        intarchivo=open(fileName+"_interfaces.json", 'a')
        dato={
            "id": self.interid,
            "name": self.intername,
            "uri": self.interuri,
            "note": self.notes,
            "available":self.interavailable,
            "deliveryMethod": self.deliveryMethod,
            "statisticsFormat": "HTML",
            "type": self.intertype
           }
        json_interfaces = json.dumps(dato)
        print('Interface: ', json_interfaces)
        intarchivo.write(json_interfaces+"\n")
    
    def printcredentials(self, idInter, login, passW, fileName):
        archivo=open(fileName+"credentials.json", 'a')
        cred ={
            "id": str(uuid.uuid4()),
            "username": login, 
            "password": passW,
            "interfaceId": idInter
             }
        json_cred = json.dumps(cred)
        print('Credentials: ', json_cred)
        archivo.write(json_cred+"\n")



#end

def divstring(stringtodivide):
    words = stringtodivide.split(' ') 
    character = ""
    for word in words:
        character += word[0]
    
    #coded=stringtodivide.split()
    character=character.upper()
    character=character.replace(")","")
    character=character.replace("(","")
    print(character)
    return character
   

class Organizations():
    #(vendorid, name, orgcod, accountingCode, language, aliases, primaryPhoneNumbers, OrgEmail, OrgUrl, contactosArray,interfacesArray)
    def __init__(self,vendorid,name,orgcode,accountingCode, language, aliases, primaryPhoneNumber, OrgEmail, OrgUrl, contactId, interfaceId):
        self.id=str(vendorid)
        self.name=name
        self.code=orgcode
        self.erpCode=accountingCode
        self.status="Active"
        self.language=language
        self.aliases=aliases
        self.addresses=[]
        self.phoneNumbers=primaryPhoneNumber
        self.emails=OrgEmail
        self.urls=OrgUrl
        self.contacts=contactId
        self.agreements=[]
        self.vendorCurrencies= ["ENG"]
        self.interfaces= interfaceId
        self.accounts=[]
        self.isVendor= True
        #self.changelogs= []


    def printorganizations(self,fileName):

        orgarchivo=open(fileName+"_organizations.json", 'a')
        organization= {
            "id": self.id,
            "name": self.name,
            "code": self.code,
            "erpCode": self.erpCode,
            "status": self.status,
            "language":self.language,
            "aliases": self.aliases,
            "addresses": self.addresses,
            "phoneNumbers": self.phoneNumbers,
            "emails": self.emails,
            "urls": self.urls,
            "contacts": self.contacts,
            "agreements": self.agreements,
            "vendorCurrencies": self.vendorCurrencies,
            "interfaces": self.interfaces,
            "accounts": self.accounts,
            "isVendor": self.isVendor,
            "paymentMethod": "EFT",
            "expectedActivationInterval": 14,
            "subscriptionInterval": 365,
            "changelogs": [],
            }
        json_organization = json.dumps(organization)
        print('Organizations: ', json_organization)
        orgarchivo.write(json_organization+"\n")
#end 


   
def interfacetype(categ):
    catego=[]
    if (categ.find('Admin') != -1):
        catego.append("Admin")
    if (categ.find('Statistics') != -1):
        catego.append("Admin")
    return catego
#end

def floatHourToTime(fh):
    h, r = divmod(fh, 1)
    m, r = divmod(r*60, 1)
    return (
        int(h),
        int(m),
        int(r*60),
    )


def cat(categ):
        dic={}
        #https://okapi-liverpool-ac-uk.folio.ebsco.com/organizations-storage/categories?query=value=="Sales"
        pathPattern="/organizations-storage/categories" #?limit=9999&query=code="
        okapi_url="https://okapi-gvsu.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJlYnNjb01pZ3JhdGlvbiIsInVzZXJfaWQiOiI2MzYzZGZmMy0yYTFhLTQ1MzgtYmRjZC1mY2Q4NTQwYjUzYzQiLCJpYXQiOjE2MDMwMzIzOTcsInRlbmFudCI6ImZzMDAwMDEwNDEifQ._26iifATVLH9jxRGkkdJxHhNWHXcNXmNli-pb-zAPgY"
        okapi_tenant="fs00001041"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="categories"
        query=f"query=value=="
        #/finance/funds?query=name==UMPROQ
        search='"'+categ+'"'
        #paging_q = f"?{query}"+search
        paging_q = f"?{query}"+search
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idcat=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idcat.append(l['id'])                   
        elif req.status_code != 400:
            idcat.append("0d8dcce9-c0c2-46fc-9957-2974bb840fff")

        return idcat
#END

def exitfile(arch):    
    if os.path.isfile(arch):
        print ("File exist")
        os.remove(arch)
    else:
        print ("File not exist")


def search(fileB,code_search):
    idlicense=""
    foundc=False
    with open(fileB,'r',encoding = 'utf-8') as h:
        for lineh in h:
            if (lineh.find(code_search) != -1):
                #print(lineh)
                foundc=True
                if (foundc):                    
                    idlicense=lineh[8:44]
                    break
    if (foundc):
        return idlicense
    else:
        idlicense="No Vendor"
        return idlicense


def make_get(pathPattern,okapi_url, okapi_tenant, okapi_token,json_file):
    pathPattern=pathPattern
    okapi_url=okapi_url
    json_file=json_file
    archivo=open(json_file+".json", 'w')
    okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
    #username="folio"
    #password="Madison"
    #payload = {'username': username, 'password': password}
    length="9999"
    start="0"
    paging_q = f"?limit={length}&offset={start}"
    path = pathPattern+paging_q
    #data=json.dumps(payload)
    url = okapi_url + path
    req = requests.get(url, headers=okapi_headers)
    if req.status_code != 201:
        print(req)
        print()
        print(req.encoding)
        print(req.text)
        print(req.headers)
        json_str = json.loads(req.text)
        #total_recs = int(json_str["totalRecords"])
        archivo.write(json.dumps(json_str, indent=2))
        print('Datos en formato JSON',json.dumps(json_str, indent=2))
        archivo.close()

def SplitString(string_to_split):
    
        #divide lastname 
        result=string_to_split.find(" ")
        #find space char
        if (string_to_split.find(' ') !=-1):
            largo=len(string_to_split)
            #print("largo:", largo)
            #print("posicion blanco:",result)
            string_fn=string_to_split[0:result]
            string_ln=string_to_split[result+1:largo]
        else:
            print("no es apellido, nombre")    
            string_fn=string_to_split
            string_ln=""
        return string_fn, string_ln
   
#end
def language_change(lan):
    idiom=""
    if lan=="English":
        idiom="eng"
    elif lan=="French":
        idiom="fre"
    elif lan=="German":
        idiom="ger"
    elif lan=="Spanish":
        idiom="spa"
        
    return idiom

def Interface_Type(itype):
    idintype=""
    if itype=="Ordering":
        idintype="Orders"
    elif itype=="Usage Reports":
        idintype="Reports"
    elif itype=="End user":
        idintype="End user"
    elif itype=="Administrative":
        idintype="Admin"
    elif itype=="Invoice":
        idintype="Invoices"
    else: 
        idintype="Other"

    return idintype

def credentials_Interface(spreadshet,wsc,Tosearch):
    
    arrayCredential=[]
    worksheetcredentials = wb.sheet_by_name(wsc)
    for l in range(worksheetcredentials.nrows):
        if l!=0:
            if worksheetcredentials.cell_value(l,0)==Tosearch:
                arrayCredential.append(str(worksheetcredentials.cell_value(l,1).strip()))
                arrayCredential.append(str(worksheetcredentials.cell_value(l,2).strip()))
    return arrayCredential
    

def readVendorspreadsheet(spreadsheet, org):
    wb = xlrd.open_workbook(spreadsheet)
    fileN=org
    worksheet = wb.sheet_by_name("vendors")
    print("no filas:", worksheet.nrows)
    print("no filas:", worksheet.ncols)
    file = open("errors.txt", "w")
    #read contacts
    sworg=0
    interfacesArray=[]
    contactosArray=[]
    oldorganizationname=""
    for p in range(worksheet.nrows):
        interfaceId=""
        contactName=""
        if p!=0:
            print("Registro No: "+str(p))
            print("Vendor: "+worksheet.cell_value(p,0))
            print("Contact People: "+worksheet.cell_value(p,8))

#############################################################################################################################
            #CONTACTS PEOPLE
#############################################################################################################################
            if worksheet.cell_value(p,12)=="":
                idcontacto=""
            else:                
                SalesContactName=str(worksheet.cell_value(p,12))
                ContactName=SplitString(SalesContactName)
                FN=ContactName[0]
                LN=ContactName[1]

                #EMAIL
                categoria=[]
                if (worksheet.cell_value(p,13)==""):
                    categoria="" #cat("General")
                else:
                    categoria=cat(worksheet.cell_value(p,13))
                
                emailsales=[]
                if worksheet.cell_value(p,14):
                    SalesContactEmail=worksheet.cell_value(p,14)
                    emailsales.append({"value": SalesContactEmail,"description": "","categories": categoria,"language": "eng"})
                
                #PHONENUMBER
                phoneNumbers=[]
                if worksheet.cell_value(p,15):
                    phoneNumbers.append({"phoneNumber": str(worksheet.cell_value(p,15)),"categories": categoria,"type": "Office","isPrimary": True,"language": "eng"})
                
                idcontacto=str(uuid.uuid4())
                org=contacts(idcontacto,FN,LN, emailsales, categoria)
                org.printcontacts(fileN,phoneNumbers)
#################CONTACT 2 ##############################################################
##SalesContactName=str(worksheet.cell_value(p,12))
            print("Contact People 2: "+worksheet.cell_value(p,16))    
            if worksheet.cell_value(p,16)=="":
                    idcontacto2=""
            else:                    
                    ContactName=SplitString(SalesContactName)
                    FN=ContactName[0]
                    LN=ContactName[1]

                #EMAIL
                    categoria=[]
                    if (worksheet.cell_value(p,17)==""):
                       categoria="" #cat("General")
                    else:
                        categoria=cat(worksheet.cell_value(p,17))
                
                    emailsales=[]
                    if worksheet.cell_value(p,18):
                        SalesContactEmail=worksheet.cell_value(p,18)
                        emailsales.append({"value": SalesContactEmail,"description": "","categories": categoria,"language": "eng"})
                    
                #PHONENUMBER
                    phoneNumbers=[]
                    if worksheet.cell_value(p,19):
                        phoneNumbers.append({"phoneNumber": str(worksheet.cell_value(p,19)),"categories": categoria,"type": "Office","isPrimary": True,"language": "eng"})
                   
                            
                    idcontacto2=str(uuid.uuid4())
                    org=contacts(idcontacto2,FN,LN, emailsales, categoria)
                    org.printcontacts(fileN,phoneNumbers)            



###############################################################################
            #INTERFACES            
###############################################################################
            print("Interfaces: "+worksheet.cell_value(p,20))
            if (worksheet.cell_value(p,20)==""):
                idInterface=""
            else:
                #Interface Type 
                if (worksheet.cell_value(p,23)):
                    InterfaceType=Interface_Type(worksheet.cell_value(p,23))
                else:
                    InterfaceType="Other"
                idInterface=str(uuid.uuid4())
                if worksheet.cell_value(p,20):
                    nameIterface= str(worksheet.cell_value(p,20).strip())
                else:
                    nameIterface="General"
                uriIterface= worksheet.cell_value(p,22)
                notesInterface= ""#worksheet.cell_value(p,7)
                availableInterface= True,
                typeInterface = [InterfaceType]
                org=interfaces(idInterface,nameIterface,uriIterface,notesInterface,availableInterface, typeInterface)
                org.printinterfaces(fileN)
                ############################################
                #CREDENTIALS
                ############################################
                interceid=worksheet.cell_value(p,21)
                if str(interceid)!="27.0" and str(interceid)!="38.0" and str(interceid)!="35.0":
                    arrayCredential=[]
                    worksheetcredentials = wb.sheet_by_name("CredentialsSourceData")
                #Search 
                    for l in range(worksheetcredentials.nrows):
                        if l!=0:
                            alex=worksheetcredentials.cell_value(l,3)
                            if worksheetcredentials.cell_value(l,3)==interceid:
                                arrayCredential.append(worksheetcredentials.cell_value(l,1))
                                arrayCredential.append(worksheetcredentials.cell_value(l,2))
                                print("Credentials: "+worksheet.cell_value(p,15))
                                org.printcredentials(idInterface, arrayCredential[0], arrayCredential[1], fileN)
                                break
                        if l=="52":
                            print(interceid)
                            file.write("Error in credentials with interface"+interceid)

                #########end credential###############################

            if (idInterface!=""):
                interfacesArray.append(idInterface)
            if idcontacto !="":
                contactosArray.append(idcontacto)
            if idcontacto2 !="":
                contactosArray.append(idcontacto)
##############################################################################################################
        #ORGANIZATIONS
############################################################################################################

            if (worksheet.cell_value(p,0)!=worksheet.cell_value(p+1,0)):
                #UUID
                vendorid=str(uuid.uuid4())
                #NAME
                name = str(worksheet.cell_value(p,0).strip())
                #CODE
                orgcod=(str(worksheet.cell_value(p,1).strip()))
                orgcod=orgcod.upper()
                #ACCOUNTING CODE
                if worksheet.cell_value(p,2):
                    accountingCode=str(worksheet.cell_value(p,2))
                    accountingCode=accountingCode.replace(".0","")
                    if len(accountingCode)!=9:
                        alen=len(accountingCode)
                        accountingCode=accountingCode.zfill(8)
                        #print(accountingCode)
                        aliases.append(worksheet.cell_value(p,8).strip())
                    else:
                        accountingCode=str(worksheet.cell_value(p,2))

                #ORGANIZATION STATUS = ACTIVE
                #LAGUAGE
                if worksheet.cell_value(p,4):
                    language=language_change(worksheet.cell_value(p,4).strip())
                #ACQUISITION UNIT
                #"acqUnitIds": 
                #DESCRIPTION
                #ALIAS    
                aliases=[]
                if worksheet.cell_value(p,8):
                        aliases.append(worksheet.cell_value(p,8).strip())


                        

                #CONTACT INFORMATION
                #ADDRESS
                #PHONENUMBER
                primaryPhoneNumbers=[]
                if worksheet.cell_value(p,9):
                    pn=str(worksheet.cell_value(p,9))
                    pn=pn.replace(".0","")
                    primaryPhoneNumbers.append({"phoneNumber": pn,"categories":cat("General"),"language":language,"type":"Other","isPrimary": False})
                #EMAIL ADDRESS
                OrgEmail=[]
                if worksheet.cell_value(p,10):
                    OrgEmail.append({"value":worksheet.cell_value(p,10).strip(),"description":"", "categories": cat("General"), "language":language})
                #URLS
                OrgUrl=[]
                if worksheet.cell_value(p,11):
                    OrgUrl.append({"value":worksheet.cell_value(p,11).strip(),"description":"","language":language, "categories": cat("General"), "notes":""})

                #ORGANIZATION OBJECT CREATION AND PRINT
                org=Organizations(vendorid, name, orgcod, accountingCode, language, aliases, primaryPhoneNumbers, OrgEmail, OrgUrl, contactosArray,interfacesArray)
                org.printorganizations(fileN)
                interfacesArray=[]
                contactosArray=[]
                print("###############################################################################################")

    file.close()
#end

if __name__ == "__main__":
    """This is the Starting point for the script"""
    customerName="gvsu"
    exitfile(customerName+"_interfaces.json")
    exitfile(customerName+"_contacts.json")
    exitfile(customerName+"_credentials.json")
    exitfile(customerName+"_organizations.json")
    readVendorspreadsheet("organizations/gvsu_vendors.xlsx", str(customerName))
    #readOrderspreadsheet(customerName+".xlsx", str(customerName))
    #exitfile(customerName+"_orders.json")
    #customerName="Liverpool"
    #readspreadsheet(customerName+".xlsx", str(customerName))

