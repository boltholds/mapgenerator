#utf-8

import sys
import openpyxl
import re
import json

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

def parse_xls(file_xls):
	wookbook = openpyxl.load_workbook(file_xls)
	worksheet = wookbook.active

	conturs = dict()

	for row in range(4, worksheet.max_row):
		iscontur = worksheet[f'C{row}'].value
		if iscontur is not None:
			conturs[iscontur] = worksheet[f'E{row}'].value
	return conturs

def print_map(map_dict):
	for name in map_dict:
		print(f"Countur {name} on adress {map_dict[name]} \n")

def zgroup(premap_dict):
	group = list()
	subgroup = list()
	prev =''
	result = premap_dict.copy()
	for glob in premap_dict:
		nl =len(glob)
		for i,elem in enumerate(premap_dict[glob]):
			if elem[2]!='':
				if len(subgroup)==0:
					loc = unpak_name(glob,elem)
					subgroup.append(loc)
					continue
				if (len(subgroup) >0) & (subgroup[0][nl] == elem[0]) & (subgroup[0][nl+2] == elem[1]):
					loc = unpak_name(glob,elem)
					subgroup.append(loc)
				else:
					name =unpak_name(glob,elem[0:2])
					#print(f"{name} : {subgroup}\n")
					if result.get(name) is None:
						result[name] = list()
					result[name] = subgroup
						
					group.append(name)
					subgroup.clear()
			elif elem[1]!='':
				name =unpak_name(glob,elem[0:2])
				group.append(name)
	
	return result

def wide_map(premap_dict):
	
	#test on relevant name
	names = r'([m]|[EKJ][Ll])[1-9]'
	sufix_pattern = r'[xX][1-9]\d*'
	match =  names+sufix_pattern
	

	global_group = dict()#
	 
	for name in premap_dict:
		res = re.fullmatch(match,name)
		#print(name if match else '')
		suf = re.findall(sufix_pattern,name)
		nm = re.findall(names,name)
		
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
		print(f"Is {key[0:2]}{main}")
		
		print(f" group {suf[0]} " if len(suf) else '')
		print(f" subgroup {suf[1]} " if len(suf)>1 else '')
		print("\n")

	'''

	
	#print_map(global_group)
	return global_group
	
def hat_countr(name,fil):
	wide = 45
	final =  f" Контур {name} ".upper()
	#print(len(final))
	#fil.write("#"*wide)
	fil.write("\n")
	fil.write("#"*10)
	fil.write(final)
	fil.write("#"*(wide-len(final)-10))
	fil.write("\n")
	#fil.write("#"*wide)
	fil.write("\n")
	
def json_parse(ispig,dirty,deep):
	#print("\t"*deep,dirty)
	print(deep)
	if type(dirty) == None:
		return 0
	elif  type(dirty) == int:
		print("\t"*len(deep),ispig," = ",dirty)
		deep.clear()
		return 0
	elif (type(dirty) == dict)&(len(dirty)>1):
		for pig in dirty:
			deep.append(pig)
			json_parse(pig,dirty.get(pig),deep)
		
			
			
			


	
def outjob(map_dir,adress_dir):
	names = {"KLall" :"клапанов", "JLall":"отсекателей","ELall":"светильников","mall":"насосов"}
	with open("map.py","w+") as out:
		out.write(hat)
		out.write("\n"*2)
		print(map_dir)
		subgroup_name = ''
		#print(map_dir)
		#print(map_dir['mall']['m3']['m3x1'])
		deep = list()
		pig =''
		json_parse(pig,map_dir,deep)

			
	

if __name__ == '__main__':

	der = parse_xls("adress.xlsx")
	outjob(wide_map(der),der)
	
	

