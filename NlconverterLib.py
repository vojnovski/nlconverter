# -*- coding: utf-8 -*-
import os 

#COM/DDE
import win32com.client #NB : Calls to COM are starting with an uppercase

#Mime dependencies
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.header
import mimetypes
from email import encoders

import re #in order to parse addresses
import tempfile #required for dealing with attachment

import datetime
import time

#mailbox
import mailbox

#Regexp
addressNotesDomainTable =  { 'dgi.finances.gouv.fr' : 'dgfip.finances.gouv.fr', }
reGenericAddressNotes = re.compile(r'CN=(.*?)\s+(.*?)\/(.*?)O=(\w*?)\/C=(\w*)', re.IGNORECASE)
reOU = re.compile(r'OU=(\w+?)\/', re.IGNORECASE)
reAddressMail = re.compile(r'([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,6})', re.IGNORECASE)
notesDll = 'C:\\Program Files (x86)\\Amanotes\\nlsxbe.dll'

def registerNotesDll():
    if os.path.exists(notesDll) and os.system('regsvr32 /s "%s"' % notesDll) == 0:
        return True
    return False

def getNotesDb(notesNsfPath, notesPasswd):
    """Connect to notes and open the nsf file"""
    session = win32com.client.Dispatch(r'Lotus.NotesSession')
    session.Initialize(notesPasswd)
    return session.GetDatabase("", notesNsfPath)


class NotesDocumentReader(object):
    """Base class for all documents"""
    def __init__(self):
        #compute a name for temporary files (attachments)
        self.tempname = os.path.join(tempfile.gettempdir(), 'nlc.tmp')

    def get(self, doc, itemname):
        """Helper to get an Item value in a document"""
        return doc.GetItemValue(itemname)

    def get1(self, doc, itemname):
        """Helper to get the first item value"""
        itemvalue = self.get(doc, itemname)
        if len(itemvalue):
            return itemvalue[0]
        else :
            return u''

    def fullDebug(self, doc):
        """Debug method : print all items values"""
        self.debugItems(doc, doc.Items)

    def debug(self, doc):
        """Debug method : print message identifiers"""
        self.debugItems(doc, ["Subject", "From", "Principal", "InetFrom", "To", "PostedDate", "DeliveredDate"])

    def debugItems(self, doc, itemlist):
        """Generic debug method"""
        self.log(20*'-')
        for it in itemlist:
            try:
                self.log("--%s = %s" % (it, doc.GetItemValue(it)) )
            except:
                self.log("--%s = !! can't display item value !!" % it)
        self.log(20*'-')

    def matchAddress(self, value):
        """Convert Notes Address Name Space into emails"""
        res = reGenericAddressNotes.search(value)
        if res == None:
            res = reAddressMail.search(value)
            if res == None:
                return value.lower()
            else :
                return res.group(1)
        else :
            mail = u"%s.%s@" % ( res.group(1).lower(), res.group(2).lower() )
            subs = reOU.findall(res.group(3))
            subs += res.groups()[3:]
            suffix = ('.'.join(subs)).lower()
            if addressNotesDomainTable.has_key(suffix):
                suffix = addressNotesDomainTable[suffix]
            mail += suffix
            return mail.lower()

    def listAttachments(self, doc):
        """Return the list of the attachments, striping None and void names"""
        return filter(lambda x : x != None and x != u'', self.get(doc, "$FILE"))

    def hasAttachments(self, doc):
        """True if theyre are any attachments"""
        return len(self.listAttachments) > 0
        
    def extractAttachment(self, doc, f):
        """Extract an attachment from the document"""
        a = doc.GetAttachment(f)

        #FIXME : bug when there is \xa0 (non breaking space) in the filename. What to do then ?
        if a == None :
            self.log("ERROR: Can't get attachment for this message :")
            self.debug(doc)
            return None
        a.ExtractFile(self.tempname)
        return self.tempname

    def dateitem2datetime(self, doc, itemname):
        datetuple = time.gmtime(int(self.get1(doc, itemname)) )[:5]
        return datetime.datetime(*datetuple )

    def log(self, message = ""):
        print message


class NotesDocumentConverter(NotesDocumentReader):
    """Base class for all converters"""

    formWhiteList = None
    formBlackList = []

    def addDocument(self, doc):
        """Generic add of a document which does nothing"""
        """Check if this form type is allowed"""
        fname = self.get1(doc, 'Form')
        return (
            (self.formWhiteList == None or fname in self.formWhiteList)
            and (fname not in self.formBlackList) #OK for blacklist
            )

    def close(self):
        pass

class NotesToMimeConverter(NotesDocumentConverter):
    """Convert a Memo Document to a Mime Message"""
    charset = 'iso-8859-15' #default charset
    charsetAttachment = 'utf-8' #attachment filename charset. Because Linux and Windows seems to use Utf-8 for filenames...
    
    def stringToHeader(self, value):
        """Build a Mail header value from a string""" 
        return email.header.Header(value, self.charset)
        
    def header(self, doc, itemname):
        return self.stringToHeader(self.get1(doc, itemname))
    
    def addressHeader(self, doc, item):
        items = self.get(doc, item)
        return self.stringToHeader(",".join(map(self.matchAddress, items)))
    
    def messageHeaders(self, doc, m):
        m['Subject'] = self.header(doc, "Subject")
        m['From'] = self.addressHeader(doc, "From")
        m['To'] = self.addressHeader(doc, "sendto")
        m['Cc'] = self.addressHeader(doc, "copyto")
        m['Date'] = self.get1(doc, "PostedDate")
        if m['Date'] == u'':
            m['Date'] = self.get1(doc, "DeliveredDate")
        ccc = self.addressHeader(doc, "BlindCopyTo")
        if ccc != u'':
            m['Ccc'] = ccc
        m['User-Agent'] = self.header(doc, "$Mailer")
        m['Message-ID'] = self.header(doc, "$MessageID")

    def buildAttachment(self, doc, f):
        """Build Mime Attachment 'f' from doc""" 
        tmp  = self.extractAttachment(doc, f)
        msg = email.mime.base.MIMEBase('application', 'octet-stream')
        if tmp != None :
            fp = open(self.tempname, 'rb')
            msg.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(msg)
            try:
                fname = f.encode(self.charsetAttachment)
            except :
                fname = f.encode(self.charset)
            msg.add_header('Content-Disposition', 'attachment', filename=fname)
        return msg

    def buildMessage(self, doc):
        """Build a message from doc"""
        main = email.mime.text.MIMEText(self.get1(doc, 'Body'), _charset=self.charset)
        
        files = self.listAttachments(doc)
        if len(files) > 0 : 
            m = email.mime.multipart.MIMEMultipart(charset=self.charset)
            m.set_charset(self.charset)
            self.messageHeaders(doc, m)
            m.attach(main)
            for it in doc.Items :
                #FIXME: should be in listAttachments
                #FIXME: should use NotesEmbedded plutôt que de scanner les items
                if it.type == 1084:
                    f = it.Values[0]
                    msg = self.buildAttachment(doc, f)
                    m.attach(msg)
        else:
            m = main
            self.messageHeaders(doc, m)
        return m

class NotesToMboxConverter(NotesToMimeConverter):
    """Notes to mbox format converter"""
    mbox = None

    def __init__(self, filename):
        super(NotesToMboxConverter, self).__init__()
        self.filename = filename
        self.mbox = mailbox.mbox(filename, None, True)
        
    def addDocument(self, doc):
        """Add a notes document to the mbox storage"""
        super(NotesToMboxConverter, self).addDocument(doc)
        m = self.buildMessage(doc)
        self.mbox.add(m)

    def close(self):
        """Close the mbox file"""
        self.log("Writing %s ... please wait." % self.filename)
        self.mbox.close()
        self.log("INFO: mbox file %s completed" % self.filename)

