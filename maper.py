#utf-8

import sys
#excel parse
import openpyxl
import xls2xlsx
import re
import json
from xls2xlsx import XLS2XLSX


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
				
def get_activesh(wookbook):
	for issheet in range(0,5):
		wookbook.active = issheet
		if wookbook.active.max_row > 4:
			return issheet
			
#read columns by adres
def parse_xls(file_xls,legacy):
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
		print(f"/!\ {file_xls} not found!")
		
	except TypeError:
		print(f"/!\ Not support type for parsing {type(file_xls)}" )
	
	else:
		start_row = init_row(worksheet)
		conturs = dict()
		for row in range(start_row, worksheet.max_row):
			iscontur = worksheet[f'C{row}'].value
			if iscontur is not None:
				conturs[iscontur] = worksheet[f'E{row}'].value
		return conturs


def print_map(map_dict):
	for name in map_dict:
		print(f"Countur {name} on adress {map_dict[name]} \n")

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
					
			
			'''
			#several print best thousand comment ;)
			print(f"Is {key[0:2]}{main}")
			
			print(f" group {suf[0]} " if len(suf) else '')
			print(f" subgroup {suf[1]} " if len(suf)>1 else '')
			print("\n")

		'''
	except TypeError:
		print(f"/!\ Instead of {type(global_group)} passed {type(premap_dict)}!")
		
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
    #print(old_lvl,lvl,keys)
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
			g =list()
			g.append('')
			rec(out,map_dir,'mall',g,-1,0)
	else:
		print("/!\  Void Dict!")


if __name__ == '__main__':

	der = parse_xls("exp/Addresses.xls",1)
	outjob(wide_map(der),"map.py")
	out_json(wide_map(der))
	

