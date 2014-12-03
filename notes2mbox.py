# -*- coding: utf-8 -*-
# hugues.bernard@gmail.com
# Pour utiliser ce script :
# * Installer python 2.6 pour windows
# * Installer pywin 2.6 pour windows
# * (optionnellement) enregistrer la dll com de notes : "regsvr32 c:\notes\nlsxbe.dll"
# * en ligne de commande (cmd) :
#   SET PATH=%PATH%;C:\Python26
#   **pour l'instant** fixer notesPasswd et notesNsfPath plus bas
#   python notes2mbox.py 
# => un fichier .mbox sera créé qu'il suffit de copier dans le répertoire ad-hoc de Thunderbird (ou d'un autre client...)

import sys
import NlconverterLib

#Constantes
notesNsfPath = 'D:\\Userfiles\\vvojnovski\\Desktop\\viktorvojnovski.nsf'
notesPasswd = 'blablabla'
outputFolder = 'D:\\Userfiles\\vvojnovski\\Desktop\\mails\\'

#Connection à Notes
db = NlconverterLib.getNotesDb(notesNsfPath, notesPasswd)

#all = tous les documents
all=db.AllDocuments
ac = all.Count
print "Nombre de documents :", ac

c = 0 #compteur de documents
e = 0 #compteur d'erreur à la conversion

# Iterates on each folder. Note that if a message is not in a folder it won't be converted
for view in db.Views:
    if view.IsFolder:
        folderName = view.Name
        folderNameClean = folderName.replace('$', '').replace('(', '').replace(')', '').replace('/', '').replace('\\', '')
        print("Processing folder %s", folderName)
        folder = db.GetView(folderName)
        if not folder:
            print('Folder "%s" not found' % folderName)
            continue
               
        mc = NlconverterLib.NotesToMboxConverter(outputFolder + folderNameClean + ".mbox")

        doc = folder.GetFirstDocument()

        while doc and c < 100000 and e < 99999:
            try:
                mc.addDocument(doc)     

            except Exception, ex:
                e += 1 #compte les exceptions
                print "\n--Exception for message %d (%s)" % (c, ex)
                mc.debug(doc)

            finally:
                doc = folder.GetNextDocument(doc)
                c+=1
                if (c % 100) == 0:
                    sys.stderr.write("%.1f%%, e=%d, c=%d\n" % (float(100.*c/ac), e, c) )
        mc.close()

print "Exceptions a traiter manuellement:", e, "... documents OK : ", c