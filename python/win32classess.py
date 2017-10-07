# -*- coding: UTF-8 -*-

import wmi, codecs, time, os, errno

# --- Dichiarazione variabili varie ---

directory_log = "logs/"         # Percorso assoluto o relativo terminante con /

delimitatore_valori = ","
delimitatori_stringa = "\""
a_capo = "\r\n"

# --- Fine dichiarazione ---

dv = unicode(delimitatore_valori)
ds = unicode(delimitatori_stringa)
ac = unicode(a_capo)
dl = directory_log

tempo_inizio = time.time()

# Inizializzo la variabile globale del CIM
cim = wmi.WMI()

# todo
def getClasse( nomeClasse, formato="csv" ):
    nomefile = "operatingSystem.csv"
    colClass = eval("cim." + nomeClasse + "()")

    if formato == "csv":
        header = unicode("")
        value = unicode("")
        for item in colClass[0].properties:
            header += ds + unicode(item) + ds + dv

        for item in colClass:
            for itemProperty in item.properties:
                s =  unicode( eval( "item." + itemProperty ) )
                value += ds + s + ds + dv
            value += ac
    else:
        if formato == "xml":
            value = unicode("")
            for item in colClass:
                tag = unicode(item.replace(" ", "_"))
                value += unicode("<") # todo
                for itemProperty in item.properties:
                    tag = unicode(itemProperty.replace(" ", "_"))
                    value += unicode("<") + tag + unicode(">")
                    value += unicode( eval( "item." + itemProperty ) )
                    value += unicode("</") + tag + unicode(">") + ac



# Creo (se non già presente) la direcotry dei log

print "Creo la struttura delle directory..."
try:
    os.makedirs(dl)
except OSError, err:
    if err.errno == errno.EEXIST:
        if os.path.isdir(dl):
            print "La directory dei log e' gia' presente, ignorato."
        else:
            print "File già esistente, ma non è directory."
            raise # re-raise the exception
    else:
        raise


# Proprietà del computer

print "Carico le risorse del computer..."
nomefile = "computerSystem.csv"
colComputer = cim.Win32_ComputerSystem()
header = unicode("")
value = unicode("")

for item in colComputer[0].properties:
    header += ds + unicode(item) + ds + dv

for item in colComputer:
    for itemProperty in item.properties:
        s =  unicode( eval( "item." + itemProperty ) )
        value += ds + s + ds + dv
    value += ac

try:
    output = codecs.open( dl + nomefile, "w", "UTF-8")
    output.write(header + ac + value)
    output.close()
except:
    print "Impossibile creare " + nomefile + "."


# Proprietà del sistema operativo

print "Carico le impostazioni di sistema..."

nomefile = "operatingSystem.csv"
colOperating = cim.Win32_OperatingSystem()
header = unicode("")
value = unicode("")

for item in colOperating[0].properties:
    header += ds + unicode(item) + ds + dv

for item in colOperating:
    for itemProperty in item.properties:
        s =  unicode( eval( "item." + itemProperty ) )
        value += ds + s + ds + dv
    value += ac

try:
    output = codecs.open( dl + nomefile, "w", "UTF-8")
    output.write(header + ac + value)
    output.close()
except:
    print "Impossibile creare " + nomefile + "."


# Proprietà dei dischi

print "Carico le impostazioni dei dischi..."

nomefile = "disk.csv"
colDisks = cim.Win32_LogicalDisk(DriveType=3)
header = unicode("")
value = unicode("")

for item in colDisks[0].properties:
    header += ds + unicode(item) + ds + dv

for item in colDisks:
    for itemProperty in item.properties:
        s = unicode( eval( "item." + itemProperty ) )
        value += ds + s + ds + dv
    value += ac

try:
    output = codecs.open( dl + nomefile, "w", "UTF-8")
    output.write(header + ac + value)
    output.close()
except:
    print "Impossibile creare " + nomefile + "."


# Proprietà di rete

print "Carico le impostazioni di rete..."

nomefile = "network.csv"
colNetwork = cim.Win32_NetworkAdapterConfiguration()
header = unicode("")
value = unicode("")

for item in colNetwork[0].properties:
    header += ds + unicode(item) + ds + dv

for item in colNetwork:
    for itemProperty in item.properties:
        s = unicode( eval( "item." + itemProperty ) )
        value += ds + s + ds + dv
    value += ac

try:
    output = codecs.open( dl + nomefile, "w", "UTF-8")
    output.write(header + ac + value)
    output.close()
except:
    print "Impossibile creare " + nomefile + "."


# Proprietà del software installato

print "Carico la lista del software..."

nomefile = "product.csv"
colProduct = cim.Win32_SoftwareFeature()
header = unicode("")
value = unicode("")

for item in colProduct[0].properties:
    header += ds + unicode(item) + ds + dv

for item in colProduct:
    for itemProperty in item.properties:
        s = unicode( eval( "item." + itemProperty ) )
        value += ds + s + ds + dv
    value += ac

try:
    output = codecs.open( dl + nomefile, "w", "UTF-8")
    output.write(header + ac + value)
    output.close()
except:
    print "Impossibile creare " + nomefile + "."

tempo_fine = time.time()

print "Fine."
print "Tempo di esecuzione: " + str(tempo_fine - tempo_inizio) + " sec."
