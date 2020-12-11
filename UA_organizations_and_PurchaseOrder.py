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

class purchaseOrder():
    def __init__(self,poNumber,vendor,orderType,notas,Order_status):
        self.poNumber=poNumber
        self.vendor=vendor
        self.orderType=orderType
        #self.notes=notas
        self.workflowStatus=Order_status
        #self.tags="EBSCOTEST"


    def printPurchaseOrderOngoingVoyager(self, Order_format, currency, renewalDate, purchase_method, eprice, id_loc, title, subscrition_from, subscription_to, package, publisher, fund, acq_method, expectedActivationDate,materialType,polId,secuence,fileName,Acqbill,manualpo,ispkg,receiptstatus,polinedescription,instructionsVendor,uuidorder,notas,instanceid):
        Ordarchivo=open(fileName+"_orders.json", 'a')
        if (Order_format=="Electronic Resource"):
            order= {
                #"id":uuidorder,
                "approved": True,
                "billTo": Acqbill,
                "manualPo": False,
                "notes": notas,
                "poNumber": self.poNumber,
                "orderType": "Ongoing",
                "reEncumber": False,
                "ongoing": {"interval": 365,"isSubscription": True,"renewalDate": renewalDate},
                "shipTo": Acqbill,
                "totalEstimatedPrice": 0.0,
                "totalItems": 1,
                "vendor": self.vendor,
                "workflowStatus": "Open",
                "compositePoLines": [
                    {
                        #"id": str(polId),
                        "checkinItems": False,
                        "acquisitionMethod": "Purchase At Vendor System",
                        "alerts": [],
                        "claims": [],
                        "collection": False,
                        "contributors": [],
                        "cost": {"listUnitPriceElectronic": 0.0,"currency": "USD","discountType": "percentage","quantityElectronic": 1,"poLineEstimatedPrice": 0.0},
                        "details": {"productIds": [],"subscriptionInterval": 0},
                        "eresource": {"activated": True,"activationDue": -623,"createInventory": "None","trial": False,"expectedActivation": expectedActivationDate,"userLimit":"","accessProvider": self.vendor,"materialType": materialType},
                        "fundDistribution": [{"code": fund[1],"fundId": fund[0],"distributionType": "percentage","value": 100.0}],
                        "isPackage": False,
                        #"locations": [{"locationId": id_loc,"quantity": 1,"quantityElectronic": 1,"quantityPhysical": 0}],
                        "orderFormat": "Electronic Resource",
                        "paymentStatus": "Awaiting Payment",
                        "physical": {"createInventory": "None","materialSupplier": self.vendor,"volumes": []},
                        "poLineNumber": self.poNumber+"-"+str(secuence),
                        "receiptStatus": "Receipt Not Required",
                        "reportingCodes": [],
                        "rush": False,
                        "source": "User",
                        "instanceId": instanceid,
                        "titleOrPackage": title,
                        "vendorDetail": {"instructions": instructionsVendor},
                        "vendorDetail": {"instructions": instructionsVendor,"refNumber": self.poNumber,"refNumberType": "Internal vendor number","vendorAccount": ""},
                     }],
                "acqUnitIds": [],
              }
        elif(Order_format=="Physical Resource"):
            order= {
                #"id":uuidorder,
                "approved": True,
                "billTo": Acqbill,
                "manualPo": False,
                "notes": notas,
                "poNumber": self.poNumber,
                "orderType": "Ongoing",
                "reEncumber": False,
                "ongoing": {"interval": 365,"isSubscription": True,"renewalDate": renewalDate},
                "shipTo": Acqbill,
                "totalEstimatedPrice": 0.0,
                "totalItems": 1,
                "vendor": self.vendor,
                "workflowStatus": "Pending",
                "compositePoLines": [
                    {
                        #"id": str(polId),
                        "checkinItems": False,
                        "acquisitionMethod": "Purchase At Vendor System",
                        "alerts": [],
                        "claims": [],
                        "collection": False,
                        "contributors": [],
                        "cost": {"listUnitPrice": 0.0,"currency": "USD","discountType": "percentage","quantityPhysical": 1,"poLineEstimatedPrice": 0.0},
                        "details": {"productIds": [],"subscriptionInterval": 0},
                        "eresource": {"activated": True,"activationDue": -623,"createInventory": "None","trial": False,"expectedActivation": expectedActivationDate,"userLimit":"",      "accessProvider": self.vendor,"materialType": materialType},
                        "fundDistribution": [{"code": fund[1],"fundId": fund[0],"distributionType": "percentage","value": 100.0}],
                        "isPackage": False,
                        "locations": [{"locationId": id_loc,"quantity": 1,"quantityElectronic": 1,"quantityPhysical": 0}],
                        "orderFormat": "Physical Resource",
                        "paymentStatus": "Awaiting Payment",
                        "physical": {"createInventory": "None","materialSupplier": self.vendor,"volumes": []},
                        "poLineNumber": self.poNumber+"-"+str(secuence),
                        "receiptStatus": "Receipt Not Required",
                        "reportingCodes": [],
                        "rush": False,
                        "source": "User",
                        "instanceId": instanceid,
                        "titleOrPackage": title,
                        #"vendorDetail": {"instructions": instructionsVendor},
                     }],
                "acqUnitIds": [],
              }
         
        json_ord = json.dumps(order,indent=2)
        json_ord = json.dumps(order)
        print('Datos en formato JSON', json_ord)
        Ordarchivo.write(json_ord+"\n")


#####==========================================================        

    def printPurchaseOrderOngoingEbscoNet(self, Order_format, currency, renewalDate, purchase_method, eprice, id_loc, title, subscrition_from, subscription_to, package, publisher, fund, acq_method, expectedActivationDate,materialType,polId,secuence,fileName,Acqbill,manualpo,ispkg,receiptstatus,polinedescription,productId,productIdType,instructionsVendor, publisher1, subscriptionInterval):
        Ordarchivo=open("Orders/"+fileName+"_orders.json", 'a')
        if (Order_format=="Electronic Resource"):
            order= {
                #"id":polId,
                "approved": True,
                "billTo": Acqbill,
                "manualPo": False,
                #"notes": "",
                "poNumber": self.poNumber,
                "orderType": "Ongoing",
                "reEncumber": True,
                "ongoing": {"interval": 365, "manualRenewal": False, "isSubscription": True,"renewalDate": renewalDate, "reviewPeriod": 90},
                "shipTo": Acqbill,
                "totalEstimatedPrice": eprice,
                "totalItems": 1,
                "vendor": self.vendor,
                "workflowStatus": "Open",
                "compositePoLines": [
                    {
                        #"id": str(polId),
                        "checkinItems": False,
                        "acquisitionMethod": "Purchase At Vendor System",
                        "alerts": [],
                        "claims": [],
                        "collection": False,
                        "contributors": [],
                        "cost": {"listUnitPriceElectronic": eprice,"currency": "USD","discountType": "percentage","quantityElectronic": 1,"poLineEstimatedPrice": eprice},
                        "details": {"productIds": [{"productId": productId,"productIdType": productIdType}],"subscriptionFrom": subscrition_from,"subscriptionTo":subscription_to,"subscriptionInterval": subscriptionInterval},
                        "eresource": {"activated": True,"activationDue": -623,"createInventory": "None","trial": False,"expectedActivation": expectedActivationDate,"userLimit":"","accessProvider": self.vendor,"materialType": materialType},
                        "fundDistribution": [{"code": fund[1],"fundId": fund[0],"distributionType": "percentage","value": 100.0}],
                        "isPackage": False,
                        #"locations": [{"locationId": id_loc,"quantity": 1,"quantityElectronic": 1,"quantityPhysical": 0}],
                        "orderFormat": "Electronic Resource",
                        "paymentStatus": "Awaiting Payment",
                        "physical": {"createInventory": "None","materialSupplier": self.vendor,"volumes": []},
                        "poLineNumber": self.poNumber+"-"+str(secuence),
                        "receiptStatus": "Receipt Not Required",
                        "reportingCodes": [],
                        "rush": False,
                        "source": "User",
                        #"instanceId": title[0],
                        "titleOrPackage": title[0],
                        "publisher": publisher1,
                        "vendorDetail": {"instructions": instructionsVendor},
                        #"vendorDetail": {"instructions": instructionsVendor,"refNumber": self.poNumber,"refNumberType": "Internal vendor number","vendorAccount": ""},
                     }],
                "acqUnitIds": [],
              }
        elif(Order_format=="Physical Resource"):
            order= {
                #"id":uuidorder,
                "approved": True,
                "billTo": Acqbill,
                "manualPo": False,
                #"notes": "",
                "poNumber": self.poNumber,
                "orderType": "Ongoing",
                "reEncumber": False,
                "ongoing": {"interval": 365,"manualRenewal": False, "isSubscription": True,"renewalDate": renewalDate, "reviewPeriod": 90},
                "shipTo": Acqbill,
                "totalEstimatedPrice": eprice,
                "totalItems": 1,
                "vendor": self.vendor,
                "workflowStatus": "Pending",
                "compositePoLines": [
                    {
                        #"id": str(polId),
                        "checkinItems": False,
                        "acquisitionMethod": "Purchase At Vendor System",
                        "alerts": [],
                        "claims": [],
                        "collection": False,
                        "contributors": [],
                        "cost": {"listUnitPrice": eprice,"currency": "USD","discountType": "percentage","quantityPhysical": 1,"poLineEstimatedPrice": eprice},
                        "details": {"productIds": [{"productId": productId,"productIdType": productIdType}],"subscriptionFrom": subscrition_from,"subscriptionTo":subscription_to,"subscriptionInterval": subscriptionInterval},
                        "eresource": {"activated": True,"activationDue": -623,"createInventory": "None","trial": False,"expectedActivation": expectedActivationDate,"userLimit":"",      "accessProvider": self.vendor,"materialType": materialType},
                        "fundDistribution": [{"code": fund[1],"fundId": fund[0],"distributionType": "percentage","value": 100.0}],
                        "isPackage": False,
                        #"locations": [{"locationId": id_loc,"quantity": 1,"quantityElectronic": 1,"quantityPhysical": 0}],
                        "orderFormat": "Physical Resource",
                        "paymentStatus": "Awaiting Payment",
                        "physical": {"createInventory": "None","materialSupplier": self.vendor,"volumes": []},
                        "poLineNumber": self.poNumber+"-"+str(secuence),
                        "receiptStatus": "Receipt Not Required",
                        "reportingCodes": [],
                        "rush": False,
                        "source": "User",
                        #"instanceId": title[0],
                        "titleOrPackage": title[0],
                         "publisher": publisher1,
                        #"vendorDetail": {"instructions": instructionsVendor},
                     }],
                "acqUnitIds": [],
              }
         
        #json_ord = json.dumps(order,indent=2)
        json_ord = json.dumps(order)
        print('Datos en formato JSON', json_ord)
        Ordarchivo.write(json_ord+"\n")

#####==========================================================      
        
    def printPurchaseOrderOnTime(self, format, currency, purchase_method,fileName):
        Ordarchivo=open(fileName+"_orders.json", 'a')
        order={

            "assignedTo": "",
            "billTo": "", 
            "shipTo": "",
            "manualPo": True,
            "approved": True, #add
            "orderType":self.orderType,
            "poNumber":self.poNumber,
            "totalItems":1,
            "vendor": self.vendor,
            "workflowStatus": "Pending",
            "notes": self.notes,
            "tags":{"tagList":[self.tags]},
            "compositePoLines": [
                {
                    #"id": alex,
                    "acquisitionMethod": purchase_method,
                    "cancellationRestriction": False,
                    "rush": False,
                    "selector": sele,
                    "cost": {"currency": "USD","listUnitPrice": eprice,"quantityPhysical": 1},
                    "locations": [{"locationId":id_loc, "quantityPhysical":1}],
                    "receiptStatus": "Awaiting Receipt",
                    "orderFormat" : "Physical Resource",
                    #"details":{"receivingNote": "ABCDEFGHIJKL"},
                    "poLineDescription": publisher,
                    "poLineNumber": polinenumber,
                    "physical":{"createInventory":"None","volumes":[vol],"materialType": mt},##add mt, exp receipt
                    "source": "User",
                    "titleOrPackage": title,
                    "fundDistribution":funds_p,#[{"code":sierra_fund_code, "fundId": fund_id, "distributionType": "percentage","value": 100}], ##add
                    "isPackage": isPkg
                 }]
                 }
        json_ord = json.dumps(order)
        print('Datos en formato JSON', json_ord)
        Ordarchivo.write(json_ord+"\n")


class contacts():

    def __init__(self,contactfirstName, contactlastName, contactemails,contactcategories ):
        self.contactid=str(uuid.uuid4())
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


    def printcontacts(self,fileName):
        contactarchivo=open(fileName+"_contacts.json", 'a')
        contacto={
            #"prefix": "",
            "id": self.contactid,
            "firstName": self.contactfirstName,
            "lastName": self.contactlastName,
            "language": self.language,
            "notes": self.contactnotes,
            "phoneNumbers": self.contactnumerotelefono,
            "emails": self.contactemails,
            "addresses": self.contactaddresses,
            "urls": self.contacturls,
            "categories": self.contactcategories,
            "inactive": self.contactinactive,
           }
        json_contact = json.dumps(contacto)
        #print('Datos en formato JSON', json_str)
        contactarchivo.write(json_contact+"\n")
        return contacto['id']
#end

class interfaces():

    def __init__(self,intername, interuri):
        self.interid=str(uuid.uuid4())
        self.intername = intername
        self.interuri = interuri
        self.deliveryMethod= "Online"
        self.interavailable=True
        self.intertype=["Admin"]
        
    def printinterfaces(self, fileName):
        intarchivo=open(fileName+"_interfaces.json", 'a')
        dato={
            "id": self.interid,
            "name": self.intername,
            "uri": self.interuri,
            "available":self.interavailable,
            "deliveryMethod": self.deliveryMethod,
            "statisticsFormat": "HTML",
            "type": self.intertype
           }
        json_interfaces = json.dumps(dato)
        #print('Datos en formato JSON', json_str)
        intarchivo.write(json_interfaces+"\n")
        return dato['id']
#end

def divstring(stringtodivide):
    coded=stringtodivide.split()
    return coded[0]
   

class Organizations():

    def __init__(self,name,contactId,interfaceId):

        self.id=str(uuid.uuid4())
        self.name=name
        self.code=divstring(name)
        self.status="Active"
        self.aliases=[]
        self.addresses=[]
        self.phoneNumbers=[]
        self.emails=[]
        self.urls=[]
        self.contacts=contactId
        self.agreements=[]
        self.vendorCurrencies= []
        self.interfaces= interfaceId
        self.accounts=[]
        self.isVendor= True
        #self.changelogs= []
    
    #methods
    def orgstatus():
        pass
    def vendor():
        pass
    def restore(self):
        pass
    def put(self,id):
        pass
    def delete(self,id):
        pass
    def str_json(self):
        pass
    def printorganizations(self,fileName):

        orgarchivo=open(fileName+"_organizations.json", 'a')
        organization= {
            "id": self.id,
            "name": self.name,
            "code": self.code,
            "status": self.status,
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
            "changelogs": [],
            }
        json_organization = json.dumps(organization)
        #print('Datos en formato JSON', json_str)
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
    catego=[]    
    if (categ.find('Renewals') != -1):
        catego.append("00a42567-e21b-4526-945e-52f2c0ed7891")#Renewals
    if (categ.find('Support') != -1):
        catego.append("1e8a8146-317a-448d-8039-fd891c0ef16e")#Support
    if (categ.find('Accounting') != -1):
        catego.append("1cdffce2-4385-49af-babe-d22df40ec207")#Accounting
    if(categ.find('General') != -1):
        catego.append("9c321c44-774c-491f-9012-2024b93cb453")#General
    if(categ.find('Sales') != -1):
       catego.append("374dfd6f-6f84-4769-82ce-0a145abeddd3")#Sales
    if(categ.find('Usage Data') != -1):
       catego.append("e45c59bb-2567-47b5-b9cf-86b3efa5364c")#usage Data
    if (categ.find('Licensing') != -1):
        catego.append("00a42567-e21b-4526-945e-52f2c0ed7891") #Renewals
    if (categ.find('Technical') != -1):
        catego.append("c885c407-f78d-4293-94ba-954ad3fc2ca8") #Renewals
    
    
    return catego
#end


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

def floatHourToTime(fh):
    h, r = divmod(fh, 1)
    m, r = divmod(r*60, 1)
    return (
        int(h),
        int(m),
        int(r*60),
    )

def is_empty(data_structure):
    if data_structure:
        print("No está vacía")
        return False
    else:
        print("Está vacía")
        return True

def getOrgId(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        pathPattern="/organizations/organizations" #?limit=9999&query=code="
        okapi_url="https://okapi-ua.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJmb2xpbyIsInVzZXJfaWQiOiJkOTE2ZTg4My1mOGYxLTQxODgtYmMxZC1mMGRjZTE1MTFiNTAiLCJpYXQiOjE1OTg1NDY2MzIsInRlbmFudCI6ImZzMDAwMDEwMDUifQ.aptR-bH8IbePZCdoGd3lomRI4-cI2jbK4AMmyAU2AOM"
        okapi_tenant="fs00001005"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"query=code=="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"?{query}"+orgname
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    idorg.append(l['name'])
        return idorg
#END
def getfunId(fund_name):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        pathPattern="/finance/funds" #?limit=9999&query=code="
        okapi_url="https://okapi-ua.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJmb2xpbyIsInVzZXJfaWQiOiJkOTE2ZTg4My1mOGYxLTQxODgtYmMxZC1mMGRjZTE1MTFiNTAiLCJpYXQiOjE1OTg1NDY2MzIsInRlbmFudCI6ImZzMDAwMDEwMDUifQ.aptR-bH8IbePZCdoGd3lomRI4-cI2jbK4AMmyAU2AOM"
        okapi_tenant="fs00001005"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="funds"
        query=f"query=name=="
        #/finance/funds?query=name==UMPROQ
        search='"'+fund_name+'"'
        #paging_q = f"?{query}"+search
        paging_q = f"?{query}"+search
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idfund=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idfund.append(l['id'])
                    idfund.append(l['name'])
        return idfund
#END

def gettitle(title_hrid):
        dic={}
        #pathPattern="/instance-storage/instances" #?limit=9999&query=code="
        #https://okapi-ua.folio.ebsco.com/instance-storage/instances?query=hrid=="264227"
        pathPattern="/instance-storage/instances" #?limit=9999&query=code="
        okapi_url="https://okapi-ua.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJmb2xpbyIsInVzZXJfaWQiOiJkOTE2ZTg4My1mOGYxLTQxODgtYmMxZC1mMGRjZTE1MTFiNTAiLCJpYXQiOjE1OTg1NDY2MzIsInRlbmFudCI6ImZzMDAwMDEwMDUifQ.aptR-bH8IbePZCdoGd3lomRI4-cI2jbK4AMmyAU2AOM"
        okapi_tenant="fs00001005"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="instances"
        query=f"query=title=="
        #/finance/funds?query=name==UMPROQ
        search='"'+title_hrid+'"'
        #paging_q = f"?{query}"+search
        paging_q = f"?{query}"+search
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idhrid=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idhrid.append(l['id'])
                    idhrid.append(l['title'])            
        return idhrid
#END


def getlocId(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        pathPattern="/locations" #?limit=9999&query=code="
        okapi_url="https://okapi-ua.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJmb2xpbyIsInVzZXJfaWQiOiJkOTE2ZTg4My1mOGYxLTQxODgtYmMxZC1mMGRjZTE1MTFiNTAiLCJpYXQiOjE1OTg1NDY2MzIsInRlbmFudCI6ImZzMDAwMDEwMDUifQ.aptR-bH8IbePZCdoGd3lomRI4-cI2jbK4AMmyAU2AOM"
        okapi_tenant="fs00001005"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="locations"
        query=f"query=name=="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"?{query}"+orgname
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    idorg.append(l['name'])
        return idorg
#END


#def EBSCO_NETreadOrderspreadsheet(spreadsheet, org):
#    wb = xlrd.open_workbook(spreadsheet)
#    fileN=org
#    worksheet = wb.sheet_by_name("AJAAeJrnlsArchivePkg")
#    print("no filas:", worksheet.nrows)
#    print("no filas:", worksheet.ncols)
    
#    #read orders
#    count=0
#    for p in range(worksheet.nrows):
#        if (p!=0):
#            if (worksheet.cell_value(p,12)!="0" and worksheet.cell_value(p,17)!=""):
#                print("Record no:"+str(count)+"Title: "+worksheet.cell_value(p,0))
#                print("Total Cost:"+str(worksheet.cell_value(p,12))+"ILS: "+str(worksheet.cell_value(p,17)))
#                #ILS NUMBER to PURCHASE ORDER NUMBER
#                orderNumber=str(worksheet.cell_value(p,17)).strip()
#                #Vendor all is EBSCO in UA https://ua.folio.ebsco.com/organizations/view/895da948-9564-57a0-a7c8-ed4c46868e6c?query=EBSCO
#                supplier="895da948-9564-57a0-a7c8-ed4c46868e6c"
#                #Order Type to Ongoing; says "renewall" default to ongoing
#                if worksheet.cell_value(p,9)=="Renewal":
#                    Ongoing="Ongoing"
#                pollinenumber=1
#                #Bill to by default Acq desk. could be change
#                billTo="ddb5ca89-0882-4557-b2ef-30f253558afa"
#                #"ongoing": {"interval": 365,"isSubscription": True,"manualRenewal": False,"reviewPeriod": 30,"renewalDate": renewalDate},

#                #TITLE
#                if (worksheet.cell_value(p,0)):
#                    title=worksheet.cell_value(p,0)
#                dt=""
#                #RenewalDate
#                if (worksheet.cell_value(p,11)):
#                    renewalDate=worksheet.cell_value(p,11)
#                    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(renewalDate) - 2)
#                    hour, minute, second = floatHourToTime(renewalDate % 1)
#                    dt = str(dt.replace(hour=hour, minute=minute,second=second))+".000+0000" #Approbal by
#                    #2019-12-12T10:11:16.449+0000
                   
#                    dt=dt.replace("2019","2021")
#                    dt=dt.replace(" ","T")
#                    renewalDate=dt
#                else:
#                    renewalDate="2021-02-10T00:00:00.000Z"

#                #quantity
#                if (worksheet.cell_value(p,7)):                      
#                    totalItems=str(worksheet.cell_value(p,7)).strip()
#                    totalItems=totalItems.replace(".0","")
#                    RenewalInterval="365"
#                    addnotas=""
#                    OrderFormat="Electronic Resource"
#                else:
#                    totalItems="1"
#                    RenewalInterval="365"
#                    addnotas=""
#                    format="Electronic Resource"

#                if (worksheet.cell_value(p,13)):
#                    currency=str(worksheet.cell_value(p,13)).strip()
#                else:
#                    currency="USD"
#                #amount
#                if (worksheet.cell_value(p,12)):
#                    amount=worksheet.cell_value(p,12)
#                else:
#                    amount="0"

#                #location by default 
#                locations="272ddb33-6f4b-4b17-92b1-6c52ed37ec82"
#                #subscritionTrom
#                dt=""
#                if (worksheet.cell_value(p,5)):
#                    subscritionFrom=worksheet.cell_value(p,5)
#                    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(subscritionFrom) - 2)
#                    hour, minute, second = floatHourToTime(subscritionFrom % 1)
#                    dt = str(dt.replace(hour=hour, minute=minute,second=second))+".000+0000" #Approbal by
#                    #2019-12-12T10:11:16.449+0000
#                    dt=dt.replace(" ","T")
#                    subscritionFrom=dt
#                #SubscriptioTo
#                dt=""
#                if (worksheet.cell_value(p,6)):
#                    subscriptionTo=worksheet.cell_value(p,6)
#                    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(subscriptionTo) - 2)
#                    hour, minute, second = floatHourToTime(subscriptionTo % 1)
#                    dt = str(dt.replace(hour=hour, minute=minute,second=second))+".000+0000" #Approbal by
#                    #2019-12-12T10:11:16.449+0000
#                    dt=dt.replace(" ","T")
#                    subscriptionTo=dt
#                #activationDate
#                expectedActivation=dt
#                #Publisher
#                if (worksheet.cell_value(p,33)):
#                    publisher=worksheet.cell_value(p,33).strip()
#                else:
#                    publisher=""
#                #Product ID
#                if (worksheet.cell_value(p,43)):
#                    publisher=worksheet.cell_value(p,43).strip()
#                else:
#                    publisher=""
#                #FUND CODE https://ua.folio.ebsco.com/finance/fund/view/77a38b62-972e-4d33-b412-e5aa780d1363?query=eng
#                fundCode="77a38b62-972e-4d33-b412-e5aa780d1363"
#                #materialType - By default serials
#                ssd=""
#                mtypes="dcbf5ae2-90d9-4cad-9cc6-13882cc2a717"
#                if (worksheet.cell_value(p,47)):
#                    uuidpol=str(worksheet.cell_value(p,47).strip()) #"5f7be7ac-6371-43a4-9a32-9f587339df66"
#                else:
#                    uuidpol="" #uuid4()
#                pollinenumber="1"

                
#                org=purchaseOrder(orderNumber,supplier,"Ongoing",addnotas,"Open")
#                #                                                                                                                             fund, acq_method, expectedActivationDate,materialType,polId,secuence,fileName):
#                #                                   OK        OK       OK                         OK          OK       OK        OK     OK                 OK            OK      OK          OK             OK              
#                org.printPurchaseOrderOngoing(OrderFormat,currency,renewalDate,"Purchase at vendor system", amount, locations, title, subscritionFrom, subscriptionTo, "True", publisher, fundCode, "Purchase At Vendor System", expectedActivation, mtypes, uuidpol, pollinenumber, fileN,billTo,"False")
###end
################################################################################################################################################################################################
################################################################################################################################################################################################

def EBSCO_NET_ACSWebEdition(spreadsheet, org):
    wb = xlrd.open_workbook(spreadsheet)
    fileN=org
    worksheet = wb.sheet_by_name("Sheet2")
    print("no filas:", worksheet.nrows)
    print("no filas:", worksheet.ncols)
    f = open("Error_Original_EBSCO_Recurring_Orders.txt", "a")
    #read orders
    count=0
    for p in range(worksheet.nrows):
        if (p!=0):
            alex= worksheet.cell_value(p,12)
            noprint=0
            if (worksheet.cell_value(p,12)!="0" and worksheet.cell_value(p,17)!=""):
                count+=1
                #prefix=""
                print("Record no:"+str(count)+"Title: "+worksheet.cell_value(p,0))
                print("Total Cost:"+str(worksheet.cell_value(p,12))+"ILS: "+str(worksheet.cell_value(p,17)))
                if "ACS Applied Materials & Interfaces"==worksheet.cell_value(p,0):
                    a=0
                #ILS NUMBER to PURCHASE ORDER NUMBER
                #PO number
                orderNumber=str(worksheet.cell_value(p,17)).strip()
                orderNumber=orderNumber.replace("-","")
                #Vendor all is EBSCO in UA https://ua.folio.ebsco.com/organizations/view/895da948-9564-57a0-a7c8-ed4c46868e6c?query=EBSCO
                supplier="895da948-9564-57a0-a7c8-ed4c46868e6c"
                vendor="895da948-9564-57a0-a7c8-ed4c46868e6c"
                #createdby
                #Created on
                #assigned to
                #Manual (Check box) Incluido en la impresion.
                #Order Type to Ongoing; says "renewall" default to ongoing
                #Reencumber (check box) yes: 
                if worksheet.cell_value(p,9):
                    Ongoing="Ongoing"
                pollinenumber=1
                #Bill to by default Acq desk. could be change
                billTo="ddb5ca89-0882-4557-b2ef-30f253558afa" #"ddb5ca89-0882-4557-b2ef-30f253558afa"
                #"ongoing": {"interval": 365,"isSubscription": True,"manualRenewal": False,"reviewPeriod": 30,"renewalDate": renewalDate},
                materialType="2e526985-5cf5-4e6b-81c9-022a68f32bec"
                mtypes=materialType
                #TITLE
                title=[]
                #if (worksheet.cell_value(p,0)):
                #    tit=str(worksheet.cell_value(p,0))
                #    title=gettitle(tit)
                #    if len(title)==0:
                #        title.append("")
                title.append(str(worksheet.cell_value(p,0)))                        
                #polinedescription=[]
                ERM_publisher= ""
                if (worksheet.cell_value(p,35)):
                    ERM_publisher= str(worksheet.cell_value(p,35))              
                    
                polinedescription=[]
                if worksheet.cell_value(p,1):
                    polinedescription.append("Title number: "+str(worksheet.cell_value(p,1)))
                if worksheet.cell_value(p,42):
                    polinedescription.append("URL: "+worksheet.cell_value(p,42))
                if worksheet.cell_value(p,4):
                    polinedescription.append("Frecuency: "+worksheet.cell_value(p,4))

                if (worksheet.cell_value(p,1)):
                    productId_ILS= str(worksheet.cell_value(p,1))
                    productIdType_ILS= "913300b2-03ed-469a-8179-c1092c991227"

                dt=""
                #RenewalDate
                if (worksheet.cell_value(p,6)):
                    
                    renewalDate=str(worksheet.cell_value(p,6))
                    if (renewalDate.find("/")>=0):
                        dt=renewalDate
                        dia=dt[0:2]
                        mes=dt[3:5]
                        ano=dt[6:10]
                        dt=ano+"-"+mes+"-"+dia+"T"+"00:00:00+0000"
                    else:
                        dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(renewalDate) - 2)
                        hour, minute, second = floatHourToTime(renewalDate % 1)
                        dt = str(dt.replace(hour=hour, minute=minute,second=second))+".000+0000" #Approbal by
                        #2019-12-12T10:11:16.449+0000
                        dt=dt.replace("2019","2021")
                        dt=dt.replace(" ","T")
                    renewalDate=dt
                else:
                    renewalDate="2021-02-10T00:00:00.000Z"

                totalEstimatedPrice="0.0"#worksheet.cell_value(p,12)
                amount="0.0"
                currency="USD"
                totalItems=str(worksheet.cell_value(p,7))
                totalItems=totalItems.replace(".0","")
                orderFormat=worksheet.cell_value(p,15)
               
                #SuscriptionFrom
                if worksheet.cell_value(p,5):
                    subscritionFrom=worksheet.cell_value(p,5)
                    if (subscritionFrom.find("/")>=0):
                        dt=subscritionFrom
                        dia=dt[0:2]
                        mes=dt[3:5]
                        ano=dt[6:10]
                        dt=ano+"-"+mes+"-"+dia+"T"+"00:00:00+0000"
                    subscritionFrom=dt
                    expectedActivationDate=dt

                dt=""
                if worksheet.cell_value(p,18):
                    renewalDate=worksheet.cell_value(p,18)
                    if (renewalDate.find("/")>=0):
                        dt=renewalDate
                        dia=dt[0:2]
                        mes=dt[3:5]
                        ano=dt[6:10]
                        dt=ano+"-"+mes+"-"+dia+"T"+"00:00:00+0000"
                    renewalDate=dt

                ERM_Subscriptioninterval=""
                if (worksheet.cell_value(p,8)):
                    ERM_Subscriptioninterval=str(worksheet.cell_value(p,8))
                    if ERM_Subscriptioninterval=="1 Year(s)":
                        ERM_Subscriptioninterval=365
                    elif ERM_Subscriptioninterval=="10 Month(s)":
                        ERM_Subscriptioninterval=304
                    elif ERM_Subscriptioninterval=="15 Month(s)":
                        ERM_Subscriptioninterval=456
                    elif ERM_Subscriptioninterval=="17 Month(s)":
                        ERM_Subscriptioninterval=517
                    elif ERM_Subscriptioninterval=="9 Month(s)": 
                        ERM_Subscriptioninterval=273


                if (orderFormat.find("/O")>=0):
                    #ELECTRONIC COST #"cost": {"listUnitPriceElectronic": totalEstimatedPrice,"currency": currency,"discountType": "percentage","quantityElectronic": 1,"poLineEstimatedPrice": totalEstimatedPrice},
                        #Cost
                    cost={"listUnitPriceElectronic": totalEstimatedPrice,"currency": currency,"discountType": "percentage","quantityElectronic": totalItems,"poLineEstimatedPrice": totalEstimatedPrice}
                    #ELECTRONIC ERESOURCE #"eresource": {"activated": True,"activationDue": -623,"createInventory": "None","trial": False,"expectedActivation": expectedActivationDate,"userLimit":"",    "accessProvider": self.vendor,"materialType": materialType},
                    eresource={"activated": True,"activationDue": -623,"createInventory": "None","trial": False,"expectedActivation": expectedActivationDate,"userLimit":"",    "accessProvider": vendor,"materialType": materialType},
                        #ELECTRONIC LOCATIONS #"locations": [{"locationId": id_loc,"quantity": 1,"quantityElectronic": 1,"quantityPhysical": 0}],
                        #locations=[{"locationId": id_loc,"quantity": 1,"quantityElectronic": 1,"quantityPhysical": 0}]

                #PHYSICAL RESOURCES
                if (orderFormat.find("/P")>=0):
                    #orderFormat=="Physical Resource"
                    #PHYSICAL "cost": {"currency": "USD","listUnitPrice": eprice,"quantityPhysical": 1},
                    cost= {"currency": currency,"listUnitPrice": totalEstimatedPrice,"quantityPhysical": totalItems},
                    #PHYSICAL "physical":{"createInventory":"None","volumes":[vol],"materialType": materialType},##add mt, exp receipt
                    physical= {"createInventory":"None","volumes":"","materialType": materialType}
                    #LOCATION PHYSICAL "locations": [{"locationId":id_loc, "quantityPhysical":1}],
                    #locations= [{"locationId":id_loc, "quantityPhysical":1}]
                    #"physical": {"createInventory": "None","materialSupplier": self.vendor,"volumes": []},
                    #physical= {"createInventory": "None","materialSupplier": vendor,"volumes": ""}
                #MIX MATERIAL
                ###if (orderFormat=="P/E MIX"):
                    #PHYSICAL "cost": {"currency": "USD","listUnitPrice": ep1rice,"quantityPhysical": 1},
                    #cost= {"currency": "USD","listUnitPrice": totalEstimatedPrice,"quantityPhysical": totalItems},
                    #PHYSICAL "physical":{"createInventory":"None","volumes":[vol],"materialType": materialType},##add mt, exp receipt
                    #physical= {"createInventory":"None","volumes":"","materialType": materialType}
                    #LOCATION PHYSICAL "locations": [{"locationId":id_loc, "quantityPhysical":1}],
                    #locations= [{"locationId":id_loc, "quantityPhysical":1}]
                    #"physical": {"createInventory": "None","materialSupplier": self.vendor,"volumes": []},
                    #physical= {"createInventory": "None","materialSupplier": vendor,"volumes": ""}

                                
                #location by default 
                #locations="272ddb33-6f4b-4b17-92b1-6c52ed37ec82"
                #subscritionTo
               
                dt=""
                if (worksheet.cell_value(p,6)):
                    subscriptionTo=worksheet.cell_value(p,6)
                    if (subscriptionTo.find("/")>=0):
                        dt=subscriptionTo
                        dia=dt[0:2]
                        mes=dt[3:5]
                        ano=dt[6:10]
                        dt=ano+"-"+mes+"-"+dia+"T"+"00:00:00+0000"
                    else:
                       
                        subscriptionTo=worksheet.cell_value(p,6)
                        dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(subscriptionTo) - 2)
                        hour, minute, second = floatHourToTime(subscriptionTo % 1)
                        dt = str(dt.replace(hour=hour, minute=minute,second=second))+".000+0000" #Approbal by
                        #2019-12-12T10:11:16.449+0000
                        dt=dt.replace(" ","T")
                    subscriptionTo=dt
                #activationDate
                expectedActivation=dt
                #Publisher
                if (worksheet.cell_value(p,33)):
                    publisher=worksheet.cell_value(p,33).strip()
                else:
                    publisher=""
                #Product ID
                if (worksheet.cell_value(p,43)):
                    publisher=worksheet.cell_value(p,43).strip()
                else:
                    publisher=""
                #FUND CODE https://ua.folio.ebsco.com/finance/fund/view/77a38b62-972e-4d33-b412-e5aa780d1363?query=eng
                
                fundname=worksheet.cell_value(p,15).strip()
                fundCode=getfunId(fundname)
                if (fundCode):
                    pass
                else:
                    noprint=1
                #materialType - By default serials
                ssd=""
                
              
                uuidpol=str(worksheet.cell_value(p,49).strip()) #"5f7be7ac-6371-43a4-9a32-9f587339df66"
                pollinenumber="1"
                addnotas=""
                OrderFormat="Electronic Resource"
                locations=""
                receiptstatus=""
                ispk= False
                if worksheet.cell_value(p,19):
                    instrVendor=worksheet.cell_value(p,19).strip()
                else:
                    instrVendor=""

                #if (noprint==0):
                org=purchaseOrder(orderNumber,vendor,"Ongoing",addnotas,"Open")
                    #countgood=+1
                    
                    #(self, Order_format, currency, renewalDate, purchase_method, eprice, id_loc, title, subscrition_from, subscription_to, package, publisher, fund, acq_method, expectedActivationDate,materialType,polId,secuence,fileName,Acqbill,manualpo,ispkg,receiptstatus,polinedescription,productId,productIdType):
                org.printPurchaseOrderOngoingEbscoNet(OrderFormat,currency,renewalDate,"Purchase At Vendor System", amount, locations, title, subscritionFrom, subscriptionTo, "True", "", fundCode, "Purchase At Vendor System", expectedActivation, mtypes, uuidpol, pollinenumber, fileN,billTo,"False",ispk,receiptstatus,polinedescription,productId_ILS,productIdType_ILS, instrVendor, ERM_publisher,ERM_Subscriptioninterval)
                
                #elif (noprint==1):
                #    countbad=+1
                #    f.write("Record #"+str(countgood)+" OrderNumber"+str(orderNumber)+"Error Funds does not exist"+str(fundname))
                #elif (noprint==2):
                #    countbad=+1
                #    f.write("Record #"+str(countgood)+" OrderNumber"+str(orderNumber)+"Error Order has hypen does not exist"+str(orderNumber))
 
##############################################################

##############################################################


def VoyagerreadOrderspreadsheet(spreadsheet, org):
    wb = xlrd.open_workbook(spreadsheet)
    fileN=org
    worksheet = wb.sheet_by_name("Voyager_Orders")
    print("no filas:", worksheet.nrows)
    print("no filas:", worksheet.ncols)
    #read orders
    count=0
    countgood=0
    countbad=0
    oldnumber=""
    addnotas=[]
    nt=0
    f = open("Error_Voyager_Orders.txt", "a")
    for p in range(worksheet.nrows):
        if (p!=0):
            #Purchase Orders
            noprint=0
            if (worksheet.cell_value(p,3)):
                orderNumber=str(worksheet.cell_value(p,3).strip())
                x=orderNumber.find('-')
                if (x!=-1):
                    orderNumber=orderNumber.replace("-","")
                    x=0
            else:
                orderNumber=""
            
            #Vendor
            if (worksheet.cell_value(p,7)):
                vendorId=str(worksheet.cell_value(p,7))
                vendorId=vendorId.replace(".0","")
                supplier=getOrgId(vendorId)
                if (supplier==""):
                    supplier=""
            if worksheet.cell_value(p,4)=="Continuation":
                    Ongoing="Ongoing"
            pollinenumber=1
                #Bill to by default Acq desk. could be change
            billTo="ddb5ca89-0882-4557-b2ef-30f253558afa"
            ##Create object 
            notasPO=[]
            if (worksheet.cell_value(p,21)):
                notasPO.append(worksheet.cell_value(p,21))
            if (worksheet.cell_value(p,23)):
                notasPO.append(worksheet.cell_value(p,23))


            print("Record #",p)
            print("Order number: ",orderNumber)
            
            

            #renewalDate
            if (worksheet.cell_value(p,0)):
                renewalDate=worksheet.cell_value(p,0)
                dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(renewalDate) - 2)
                hour, minute, second = floatHourToTime(renewalDate % 1)
                dt = str(dt.replace(hour=hour, minute=minute,second=second))+".000+0000" #Approbal by
                dt=dt.replace("2019","2021")
                dt=dt.replace(" ","T")
                renewalDate=dt
            #quantity
            if (worksheet.cell_value(p,15)):                      
                totalItems=str(worksheet.cell_value(p,15)).strip()
                totalItems=totalItems.replace(".0","")
                RenewalInterval="365"
                addnotas=""
                OrderFormat="Electronic Resource"
            else:
                totalItems="1"
                RenewalInterval="365"
                addnotas=""
                OrderFormat="Electronic Resource"
            #is a package
            ispk="False"
            titleindex=[]
            instanceid=""
            title=""
#            #Title
            if (worksheet.cell_value(p,2)):
                hrid=str(worksheet.cell_value(p,2))
                hrid=hrid.replace(".0","")
                titleindex=gettitle(hrid)
                if is_empty(titleindex):
                    title=worksheet.cell_value(p,14).strip()
                    instanceid=""
                    nt=nt+1
                else:
                    instanceid=str(titleindex[0])
                    print(instanceid)
                    title=titleindex[1]
            
            subscritionFrom=""
            subscriptionTo=""

            if (worksheet.cell_value(p,18).strip()):
                of=worksheet.cell_value(p,18)
                if (of.find("/P")!=-1):
                    OrderFormat="Physical Resource"
                else:
                    OrderFormat="Electronic Resource"

            if (worksheet.cell_value(p,13)):
                rs=worksheet.cell_value(p,13)
                if (rs=="Approved"):
                    receiptstatus="Fully Received"
                elif (rs=="Received Partial"):
                    receiptstatus="Partially received"
                else:
                    receiptstatus="Fully Received"
            
            if (worksheet.cell_value(p,20)):
                polinedescription=""
                polinedescription=worksheet.cell_value(p,20).strip()
            #Price 0
            amount="0"
            #USD by default
            currency="USD"
            #all were assigned to mus/ff, UA need to create funds according to
            fundname= worksheet.cell_value(p,18)
            fundCode=getfunId(fundname)
            if (fundCode):
                pass
            else:
                noprint=1
            
            mtypes="2e526985-5cf5-4e6b-81c9-022a68f32bec"
            locations=str("624eaa3f-2020-45e0-a064-f3d79b31a094")
            if (worksheet.cell_value(p,16)):
                if worksheet.cell_value(p,16)=="Sci and Eng Library":
                    locations=str("624eaa3f-2020-45e0-a064-f3d79b31a094")
                elif worksheet.cell_value(p,16)=="Electronic Book":
                    #Request the true locations
                    locations=str("624eaa3f-2020-45e0-a064-f3d79b31a094")
                else: 
                    orgname=worksheet.cell_value(p,16)
                    locations=getlocId(orgname)

            expectedActivation=subscriptionTo
            if (worksheet.cell_value(p,20)):
                instructionsVend=""
                instructionsVend=str(worksheet.cell_value(p,20).strip())
            else:
                instructionsVend=""
            #Purchase Order
            uuidpol=str(worksheet.cell_value(p,25))
            #Order
            uuidorder=str(worksheet.cell_value(p,26))

            
            if (noprint==0):
                org=purchaseOrder(orderNumber,supplier[0],"Ongoing",addnotas,"Open")
                #countgood=+1
                #B.write("Record #"+str(countgood)+" OrderNumber"+str(orderNumber))
                if (oldnumber==orderNumber):
                    org.printPurchaseOrderOngoingVoyager(OrderFormat,currency,renewalDate,"Purchase At Vendor System", amount, locations[0], title, ubscritionFrom, subscriptionTo, "True", "", fundCode, "Purchase At Vendor System", expectedActivation, mtypes, uuidpol, ollinenumber, fileN,billTo,"False",ispk,receiptstatus,polinedescription,instructionsVend,uuidorder,notasPO,instanceid)
                else:
                    org.printPurchaseOrderOngoingVoyager(OrderFormat,currency,renewalDate,"Purchase At Vendor System", amount, locations[0], title, subscritionFrom, subscriptionTo, "True", "", fundCode, "Purchase At Vendor System", expectedActivation, mtypes, uuidpol, pollinenumber, fileN,billTo,"False",ispk,receiptstatus,polinedescription,instructionsVend,uuidorder,notasPO,instanceid)
                oldnumber=orderNumber
            elif (noprint==1):
                countbad=+1
                f.write("Record #"+str(p)+" OrderNumber"+str(orderNumber)+"Error Funds does not exist"+str(fundname)+"\n")
            elif (noprint==2):
                countbad=+1
                f.write("Record #"+str(p)+" OrderNumber"+str(orderNumber)+"Error Order has hypen does not exist"+str(orderNumber)+"\n")
    print("bad records:",str(countbad))
    print("not titles: ",nt)
            
##end

if __name__ == "__main__":
    """This is the Starting point for the script"""
    #EBSCO_NETreadOrderspreadsheet(customerName+".xlsx", "EbscoNET_to_folio_orders")
    sw=1
    customerName="EBSCOnet_to_folio_orders"
    #customerName="BR16245"
    if (sw==1):
        VoyagerreadOrderspreadsheet(customerName+".xlsx", "Voyager_to_FOLIO_orders")
    elif(sw==2):
        EBSCO_NET_ACSWebEdition("Orders/"+customerName+".xlsx", "EBSCO_NET")
    elif(sw==3):
        pass

