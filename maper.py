#utf-8

import sys
#excel parse
import openpyxl
import xls2xlsx
import re
import json
from collections import deque
from xls2xlsx import XLS2XLSX
import logging

#logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(name)s %(levelname)s:%(message)s')
logging.getLogger(__name__)

hat = '''#utf-8
module0	= [] #default
module1	= []
module2	= []
module3	= []

##############################################################
delay = 100 # sound delay ####################################
hz = 50 # max value in hz ####################################
bar = 16 # max value in bar ##################################
colorCodingMode = 1 # how many colors one byte encodes #######
minBrightnessRed = 128 # for colorCodingMode 2 or 3 ########## in developing
minBrightnessGreen = 128 # for colorCodingMode 2 or 3 ######## in developing
minBrightnessBlue = 128 # for colorCodingMode 2 or 3 ######### in developing
##############################################################
'''
names = {"KLall" :"клапанов", "JLall":"отсекателей","ELall":"светильников","mall":"насосов"}


def init_row(worksheet):
	for row in range(1, worksheet.max_row):
		rowvalue = worksheet[f'B{row}'].value
		if rowvalue:
			if rowvalue[:3] == "ARK":
				return row
	logging.error("No named ARK in address file!")
				
def get_activesh(wookbook):
	for issheet in range(0,5):
		wookbook.active = issheet
		if wookbook.active.max_row > 4:
			return issheet
			
#read columns by adres
def parse_xls(file_xls,legacy):
	#offset_transmiter = 512
	try:
		if not legacy:
			wookbook = openpyxl.load_workbook(file_xls, data_only=True)
			
			#found start row of adress name
		else:
			w2x = XLS2XLSX(file_xls)
			wookbook = w2x.to_xlsx()
		wookbook.active = get_activesh(wookbook)
		worksheet = wookbook.active

		
	except FileNotFoundError:
		logging.error(f" {file_xls} not found!")
		
	except TypeError:
		logging.error(f" Not support type for parsing {type(file_xls)}" )
	
	else:
		start_row = init_row(worksheet)
		conturs = dict()
		for row in range(start_row, worksheet.max_row):
			iscontur = worksheet[f'C{row}'].value
			try:
				val = worksheet[f'E{row}'].value
				if iscontur is not None:
					if conturs.get(iscontur) is None:
						conturs[iscontur] = val
					else:
						raise (AssertionError)
			except AssertionError:
				logging.warning(f" Counture {iscontur} = [ {val} ] has a duplicate! Was remove")
				
		return conturs



def print_map(map_dict):
	for name in map_dict:
		logging.info(f" Countur {name} on adress {map_dict[name]} \n")

#parse the hierarchy counturs on map to the JSON(dict in dict)
def wide_map(premap_dict):
	
	#NAME is ELx,mx,JLx
	names = r'([m]|[EKJ][Ll])[1-9]'
	#it's xNxN exp x1 or x2x1
	sufix_pattern = r'[xX][1-9]\d*'
	#NAMExNxN,NAMExN
	match =  names+sufix_pattern
	

	global_group = dict()#result
	try:
		for name in premap_dict:
			res = re.fullmatch(match,name)
			suf = re.findall(sufix_pattern,name)
			nm = re.findall(names,name)
			if nm:
				glob = f"{nm[0]}"
				group = suf[0][1:] if len(suf) else ''
				subgroup = suf[1][1:] if len(suf)>1 else ''
				main =  name[len(nm[0])]
				key=glob+'all'
				if global_group.get(key) is None:
					global_group[key] = dict()
				counter_name = glob+main
				if global_group[key].get(counter_name) is None:
					global_group[key][counter_name] = dict()

				if(group != ''):
					subcounter_name = counter_name+'x'+group
					if	(global_group[key][counter_name].get(subcounter_name) is None):
						global_group[key][counter_name][subcounter_name] =dict()
					if subgroup != '':
						element = subcounter_name+'x'+subgroup
						global_group[key][counter_name][subcounter_name][element] = premap_dict[name]
					else:
						global_group[key][counter_name][subcounter_name] = premap_dict[name]
				else:
					global_group[key][counter_name] = premap_dict[name]
					
			
		
			#several logging best thousand comment ;)
			logging.debug(f" Is {key[0:2]}{main}")
			
			logging.debug(f" group {suf[0]} " if len(suf) else '')
			logging.debug(f" subgroup {suf[1]} " if len(suf)>1 else '')
			logging.debug("\n")


	except TypeError:
		logging.error(f" Instead of {type(global_group)} passed {type(premap_dict)}!")
		
	return global_group

#splitter from countur group	
def hat_countr(name,fil):
	wide = 45
	final =  f" Контур {name} ".upper()
	fil.write("\n")
	fil.write("#"*10)
	fil.write(final)
	fil.write("#"*(wide-len(final)-10))
	fil.write("\n")
	fil.write("\n")


			
#dump file		
def out_json(main_dir):
	with open("map.json","w+") as out:
		js =json.dumps(main_dir, indent=4)
		out.write(js)

#recursive walk on tree heararhy
#if we down to deep, then add eque name node
#if we up from deep ,the pop eaue name node
def rec(outfile,mapy,keys,group,lvl,old_lvl):
    old_lvl = lvl
    lvl+=1
    if (names.get(keys) != None)&(lvl<=0):
    	#if on head then print the hat :)
    	hat_countr(names.get(keys),outfile)
    for key,value in mapy.items(): 
        #if it's terminal
        if type(value) == int:
            outfile.write(f"{key} = [ {value} ]\n")
		#if it vertex
        elif type(value) == dict:
            
            group.append(f"{key} = " +" + ".join(list(value.keys()))+"\n ") 
            old_lvl = lvl
            lvl+=rec(outfile,value,key,group,lvl,old_lvl)
    #print countur hierarhy
    outfile.write("\n")
    outfile.write(group.pop()) 
    outfile.write("\n")  
    return -1 #up level on JSON


#write to map.py
def outjob(map_dir,filename):
	if map_dir:
		with open(filename,"w+") as out:
			out.write(hat)
			out.write("\n"*2)
			subgroup_name = ''
			g = deque()
			g.append('')
			rec(out,map_dir,'mall',g,-1,0)
	else:
		logging.error(" Void passed dict")


if __name__ == '__main__':

	der = parse_xls("exp/adress_dubleerrtets.xlsx",0)
	outjob(wide_map(der),"map.py")
	out_json(wide_map(der))
	

