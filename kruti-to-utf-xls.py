import xlrd
import xlwt
from sys import argv,exit
from kruti_to_uni8 import kruti_to_utf8 as kd2u8
from platform import platform

if (platform()[0:7]=="Windows"):
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

class bcolors:
    HEADER = '\033[95m'
    OKGREEN = '\033[92m'
    OKBLUE = '\033[94m'
    OKPURPLE = '\033[95m'
    INFOYELLOW = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

def sane_write(inputval):
    if(isfloat(inputval)):
        if(float(inputval)-int(float(inputval))==0):
            write_conv=kd2u8(str(int(inputval)))
        else:
            write_conv=kd2u8(str(inputval))
    else:
        write_conv=kd2u8(inputval)
    return write_conv

def processxls(inputfile,outputfile):
    wb2=xlwt.Workbook()
    wb=xlrd.open_workbook(inputfile)
    print(f'\nLoaded:{bcolors.OKBLUE}{inputfile}{bcolors.ENDC}')
    sheetnames=wb.sheet_names()
    smax=len(sheetnames)
    s=0
    for names in sheetnames:
        s=s+1
        newsheet=wb2.add_sheet(names)
        sh=wb.sheet_by_name(names)
        rmax=sh.nrows-1
        cmax=sh.ncols-1
        print(f'Processing Sheet:{bcolors.OKBLUE}{sh.name}{bcolors.ENDC},ROWS:{bcolors.OKBLUE}{sh.nrows}{bcolors.ENDC},COLUMNS:{bcolors.OKBLUE}{sh.ncols}{bcolors.ENDC}\n')

        for i in range(rmax):
            pvalue=round(i/(rmax-1)*100,2)
            overall_val=round(100.00*(s-1)/smax+pvalue/smax,2)
            ins=''
            for k in range(int(pvalue/5)):
                ins=ins+'#'
                if(k==19):
                    ins=ins+'\033[1D#]'
            print(f'\033[A [{bcolors.OKGREEN}{overall_val:.2f}%{bcolors.ENDC}][{bcolors.OKBLUE}{pvalue:.2f}{bcolors.ENDC}%][{ins}{bcolors.OKBLUE}#{bcolors.ENDC}  ')
            for j in range(cmax):
                write_conv=sane_write(sh.cell_value(i,j))
                newsheet.write(i,j,write_conv)
        print(f'Writing:{bcolors.OKBLUE}{outputfile}{bcolors.ENDC}\n')
        wb2.save(outputfile)

try:
    processxls(argv[1],argv[2])
except Exception as e:
    print(e)