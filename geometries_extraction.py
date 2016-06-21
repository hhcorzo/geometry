import os
import re
import openpyxl
workbook = openpyxl.Workbook()  #creates openpyxl workbook

#logFilesFolder is name of folder containing log files
logFilesFolder='/Geom_for_code'

'''path to this file'''
path=os.path.dirname(os.path.realpath(__file__))
pathorigin=path     #used to save workbook in this location
excelFilePathName='/geometry_logFile_data.xlsx'

#columns for each variable in workbook
colFileInformation='A'
colMolecule='B'
colCharge='C'
colMultiplicity='D'
colBasis='E'
colSymmetry='F'
colHarmonicFrequency='G'

#prepare openpyxl first
worksheet=workbook.active
worksheet.title="Data"
#creates worksheet Data
#add headings to each column
worksheet[colFileInformation+'1']='File'
worksheet[colMolecule+'1']='Molecule'
worksheet[colCharge+'1']='Charge'
worksheet[colMultiplicity+'1']='Multiplicity'
worksheet[colBasis+'1']='Basis'
worksheet[colSymmetry+'1']='Symmetry'
worksheet[colHarmonicFrequency+'1']='Harmonic Frequency'

def writeDataToExcel(row,fileInformation,molecule,charge,multiplicity,basis,symmetry,harmonicFrequency):
    '''writesDataToExcel takes is called by dataExtract. It takes in the variables found in
    data extraction and writes it into the openpyxl workbook'''

    worksheet[colFileInformation+str(row)]=fileInformation
    worksheet[colMolecule+str(row)]=molecule
    worksheet[colCharge+str(row)]=charge
    worksheet[colMultiplicity+str(row)]=multiplicity
    worksheet[colBasis+str(row)]=basis
    worksheet[colSymmetry+str(row)]=symmetry
    worksheet[colHarmonicFrequency+str(row)]=harmonicFrequency

def dataExtract(path):

    row=2
    #extraction from log files starts here
    logFiles=[]

    for path, subdirs, files in os.walk(path+logFilesFolder):
        for name in files:
            if os.path.join(path, name)[len(os.path.join(path, name))-4:len(os.path.join(path, name))]=='.log':
                logFiles.append(os.path.join(path, name))

    for currentFile in logFiles:
        log = open(currentFile, 'r').read()
        splitLog = re.split(r'[\\\s]\s*', log)  #splits string with \ (\\), empty space (\s) and = and ,
        x=0
        while x<len(splitLog):
            if splitLog[x]=='Stoichiometry':
                molecule=splitLog[x+1]
            if splitLog[x]=='Charge' and splitLog[x-1]=='Z-matrix:':
                charge=splitLog[x+2]
            if splitLog[x]=='Multiplicity':
                multiplicity=splitLog[x+2]
            if splitLog[x]=='Standard' and splitLog[x+1]=='basis:':
                basis=splitLog[x+2] +' '+splitLog[x+3]+splitLog[x+4]
            if splitLog[x]=='Full' and splitLog[x+1]=='point' and splitLog[x+2]=='group':
                symmetry=splitLog[x+3]
            if splitLog[x]=='normal' and splitLog[x+1]=='coordinates:':
                y=0
                while splitLog[x+y]!='Frequencies':
                    y+=1
                harmonicFrequency=splitLog[x+y+2]

            x+=1
        fileInformation=currentFile

        writeDataToExcel(row,fileInformation,molecule,charge,multiplicity,basis,symmetry,harmonicFrequency)
        row+=1
    workbook.save(pathorigin + excelFilePathName)     #saves file


def run():
    dataExtract(path)
    
run()    