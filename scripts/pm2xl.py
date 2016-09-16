
qshome='C:/PortraitMiner7.1/server/qs7.1/win64/bin/'

import os
import pandas as pd
import xlsxwriter
import argparse

## USAGE: pm2xl.py -f ../data/ConsentSegmentStats.ftr -x ../temp/MyExcel.xlsx -s foo,bar -r 1,rownum()<10
## Spaces in arguments don't work!

parser = argparse.ArgumentParser(description='This is a Miner to XLSX utility template')
#pm,xl,sheets,selections=None,fields=None,index=None
parser.add_argument('-f','--focus', help='full focus name',required=True)
parser.add_argument('-x','--excel',help='full Excel filename', required=True)
parser.add_argument('-s','--sheets',help='comma list of sheet names', required=True)
parser.add_argument('-r','--selections',help='comma list of fdl for record selection', required=True)
args = parser.parse_args()




def runqsdb(command, args, failonbad=True):

    print('EXECUTING',qshome+command,[qshome+command]+args)
    result = os.spawnv(os.P_WAIT, qshome+command, [qshome+command]+args)
    if result==1 and failonbad:
            raise Exception( qshome+command,' failed for ',[command]+args)
    return result


def qsexportflat(input, output, records=None, fields=None, headers=True):

    args=[]
    args.extend(['-input',input,'-output',output])

    for arg in ['records','fields']:
        if eval(arg):
            args.extend(['-'+arg,eval(arg)])

    for arg in ['headers']:
        if eval(arg):
            args.extend(['-'+arg])

    runqsdb("qsexportflat.exe", args)


def pm2xl(pm,xl,sheets,selections=None,fields=None,index=None):
    # Use index only if NOT first column
    tempff='../temp/ff.txt'
    print(pm,xl,sheets,selections)
    writer = pd.ExcelWriter(xl, engine='xlsxwriter')

    for i in range(0,len(sheets)):
        sheet=sheets[i]
        selection=selections[i]
        print(i,sheet,selection,index)
        qsexportflat(pm,tempff,records=selection,fields=fields)

        if index==None:
            df = pd.read_csv(tempff);
        else:
            df = pd.read_csv(tempff,index_col=index);

        df.to_excel(writer, sheet_name=sheet)

    writer.save()


f='../data/ConsentSegmentStats'
s='BettingSegment="""All""" '
p='"rownum() < 12"'
e='../temp/Excel.xlsx'

#pm2xl(f,e,['foo'],[s])

lsheets = [str(item) for item in args.sheets.split(',')]
lselections = [str(item) for item in args.selections.split(',')]
print(lsheets,lselections)
pm2xl(args.focus,args.excel,lsheets,lselections)
