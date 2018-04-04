import openpyxl
import pyperclip
import re
import sys
import os
import datetime
from pprint import pprint



def read_file() -> dict:
	
	print('Имя файла или Enter(input.xlsx - по умолчанию)')
	file_name = input()
	if file_name == '':
		wb = openpyxl.load_workbook('inputt.xlsx')
	else:
		wb = openpyxl.load_workbook(file_name)
	
	
	print('Название листа или Enter(Лист1 - по умолчанию): ')
	sheet_name = input()
	
	if sheet_name == '':
		sheet = wb['Лист1']
	else:
		sheet = wb[str(sheet_name)]
	
	sheet_dict = {'sheet_obj': sheet,'max_row': sheet.max_row, 'max_column': sheet.max_column}
	
	return sheet_dict
	
def compile_table() -> dict:
	sheet = read_file()
	header_alias = alias_dict()
	table_dict = {}
	
	for y in range(1,int(sheet.get('max_column')) + 1):
		for k, v in header_alias.items():
			regex = re.compile(v, re.IGNORECASE)
			reg_res = re.search(regex, sheet.get('sheet_obj').cell(row=1, column=y).value)
			if reg_res:
				table_header = str(reg_res.group(2)) + str(reg_res.group(3))
				table_header = table_header.upper()
				print(str(y) + ' Regex result: ' + str(reg_res.group(2)) + str(reg_res.group(3)))
				row_value_lst = []
				for i in range(2,int(sheet.get('max_row'))+1):
						temp = str(sheet.get('sheet_obj').cell(row=i, column=y).value)
						row_value_lst += [temp]
						table_dict[table_header] = row_value_lst
					
			#else:
				#print(str(y) + 'No match')

	table_dict['len_of_element_list'] = sheet.get('max_row') - 1
	return table_dict.copy()	

def alias_dict() -> dict:
	
	# header_alias_t = {}
	# with open('tr_templates.txt') as rf:
		# for line in rf:
			# (key, val) = line.split()
			# header_alias_t[key] = val
 	
	header_alias = {'CELLNAME':				'((^CELL)[-_\s]*?(NAME$))',\
					'TRXID':				'((TRX)[-_\s]*?(ID))',\
					'TRXNAME':				'((TRX)[-_\s]*?(NAME))',\
					'EGBTSPOWT':			'((EGBTS)[-_\s\S]*?(POWT))',\
					'RNCID':				'((\bRNC)[-_\s\S]*?(ID\b))',\
					'DRDECN0THRESHHOLD':	'((^DRDECN0))[-_\s\S]*?(THRESHHOLD)',\
					'VENDORSRC': 			'((VENDOR)[-_\s]*?(SRC))',\
					'NRNCID':				'((^NRNC)[-_\s\S]*?(ID\b))',\
					'RNCID':				'((^RNC)[-_\s\S]*?(ID$))',\
					'CELLID':				'((^CELL)[-_\s]*?(ID$))',\
					'BANDIND':				'((BAND)[-_\s]*?(IND))',\
					'UARFCNUPLINKIND':		'((UARFCN)[-_\s]*?(UPLINKID))',\
					'UARFCNUPLINK':			'((UARFCN)[-_\s]*?(UPLINK))',\
					'UARFCNDOWNLINK':		'((UARFCN)[-_\s]*?(DOWNLINK))',\
					'NCELLRNCID':			'((^NCELLRNC)[-_\s]*?(ID))',\
					'NCELLID':				'((^NCELL)[-_\s]*?(ID))',\
					'BSC Name Src':			'((BSCNAME)[-_\s]*?(SRC$))',\
					'GEXT2GCELL':			'((G?EXT2G)[-_\s]*?(CELL))',\
					'VALUEDST':				'((VALUE)[-_\s]*?(DST$))',\
					'PARAMETER':			'((PARA)[-_\s]*?(METER))',\
					'PATHID':				'((PATH)[-_\s]*?(ID$))',\
					'ANI':					'((A)[-_\s]*?(NI))',\
					'NODEB_NAME':			'((^NODEB)[-_\s]*?(NAME$))',\
					'FREQ1':				'((^FRE)[-_\s]*?(Q1$))',\
					'FREQ':					'((^FRE)[-_\s]*?(Q$))',\
					'BLINDHOFLAG':			'((^BLINDHOF)[-_\s]*?(LAG$))',\
					'BLINDHOQUALITY':		'((^BLINDHOQUALITY)[-_\s]*?(CONDITION$))',\
					'EARFCN':				'((^EAR)[-_\s]*?(FCN$))',\
					'NPRIORITY':			'((^NPRIO)[-_\s]*?(RITY$))',\
					'THDTOLOW':				'((^THDTO)[-_\s]*?(LOW$))',\
					'THDTOHIGH':			'((^THDTO)[-_\s]*?(HIGH$))',\
					'EMEASBW':				'((^EMEA)[-_\s]*?(SBW$))',\
					'EQRXLEVMIN':			'((^EQRXLEV)[-_\s]*?(MIN$))',\
					'EDETECTIND':			'((^EDETECT)[-_\s]*?(IND$))',\
					'SUPCNOPGRPINDEX':		'((^SUPCNO)[-_\s]*?(PGRPINDEX$))',\
					'BLACKLSTCELLNUMBER':	'((^BLACKLSTCELL)[-_\s]*?(NUMBER$))',\
					'RSRQSWITCH':			'((^RSRQ)[-_\s]*?(SWITCH$))',\
					'NPRIORITYCONNECT':		'((^NPRIORITY)[-_\s]*?(CONNECT$))',\
					'EQRXLEVMINOFFSET':		'((^EQRXLEVMIN)[-_\s]*?(OFFSET$))',\
					'EQQUALMINOFFSET':		'((^EQQUALMIN)[-_\s]*?(OFFSET$))',\
					'EQRXLEVMINSTEP':		'((^EQRXLEVMIN)[-_\s]*?(STEP$))',\
					'EQQUALMINSTEP':		'((^EQQUALMIN)[-_\s]*?(STEP$))',
					'NODEBID':				'((^NODEB)[-_\s]*?(ID$))',\
					'SRC3GNCELLNAME':		'((^SRC3GNCELLN)[-_\s]*?(AME$))',\
					'SRC2GNCELLID':			'((^SRC2GNCELL)[-_\s]*?(ID$))',\
					'NBR3GNCELLNAME':		'((^NBR3GN)[-_\s]*?(CELLNAME$))',\
					'SRC3GNCELLID':			'((^SRC3GNCELL)[-_\s]*?(ID$))',\
					'NBR3GNCELLID':			'((^NBR3GNCELL)[-_\s]*?(ID$))',\
					'NBR2GNCELLID':			'((^NBR2GNCELL)[-_\s]*?(ID$))',\
					'T3168':				'((^T)[-_\s]*?(3168$))',\
					'INBSCHOTIMER':			'((^INBSCHO)[-_\s]*?(TIMER$))',\
					'WAITFORRELIND':		'((^WAITFOR)[-_\s]*?(RELIND$))',\
					'IMMREJWAITINDTIMER':	'((^IMMREJWAITIND)[-_\s]*?(TIMER$))',\
					'TIQUEUINGTIMER':		'((^TIQUEUING)[-_\s]*?(TIMER$))',\
					'T200SDCCH':			'((^T200)[-_\s]*?(SDCCH$))',\
					'T200FACCHF':			'((^T200)[-_\s]*?(FACCHF$))',\
					'T200FACCHH':			'((^T200)[-_\s]*?(FACCHH$))',\
					'T200SACCT0':			'((^T200)[-_\s]*?(SACCT0$))',\
					'T200SACCH3':			'((^T200)[-_\s]*?(SACCH3$))',\
					'T200SACCHS':			'((^T200)[-_\s]*?(SACCHS$))',\
					'T200SDCCH3':			'((^T200)[-_\s]*?(SDCCH3$))',\
					'RA':					'((^R)[-_\s]*?(A$))',\
					'GSMCELLINDEX':			'((^GSM)[-_\s]*?(CELLINDEX$))',\
					'LA':					'((^LA)[-_\s]*?(C$))',\
					'SRCCI':				'((^SRC)[-_\s]*?(CI$))',\
					'SEARCHVALUE':			'((^SEARCH)[-_\s]*?(INDEX$))',\
					'VALUELIST':			'((^VALUE)[-_\s]*?(LIST$))',\
					'RESULTVALUE':			'((^RESULT)[-_\s]*?(VALUE$))',\
					'MAXTXPOWER':			'((^MAXTX)[-_\s]*?(POWER$))',\
					'SRCCELLNAME':			'((^SRC)[-_\s]*?(CELLNAME$))',\
					'NBRCELLNAME':			'((^NBR)[-_\s]*?(CELLNAME$))',\
					'SRC3GNCELLNAME':		'((^SRC3GN)[-_\s]*?(CELLNAME$))',\
					'NBR3GNCELLNAME':		'((^NBR3GN)[-_\s]*?(CELLNAME$))',\
					'LOCALCELLID':			'((^LOCALCELL)[-_\s]*?(ID$))',\
					'CellReselPriority':	'((^CellResel)[-_\s]*?(Priority$))',\
					'ThreshXhigh':			'((^Thresh)[-_\s]*?(Xhigh$))',\
					'ThreshXlow':			'((^Thresh)[-_\s]*?(Xlow$))',\
					'ConnFreqPriority':		'((^ConnFreq)[-_\s]*?(Priority$))',\
					'MeasFreqPriority':		'((^MeasFreq)[-_\s]*?(Priority$))',\
					'PCPICHPOWER':			'((^PCPICH)[-_\s]*?(POWER$))'}
	
	return header_alias
	

def generate_script():
	#input_string = 'ADD UINTERFREQNCELL:RNCID=169,CELLID=2141,NCELLRNCID=12312,NCELLID=234,SIB11IND=TRUE,SIB12IND=FALSE,TPENALTYHCSRESELECT=D0,BLINDHOFLAG=FALSE,NPRIOFLAG=FALSE,INTERNCELLQUALREQFLAG=FALSE,CLBFLAG=FALSE;'
	
	table_dict = compile_table()
	input_string = get_string()
	element_length = table_dict['len_of_element_list']
	final_script_list = []
	index_of_final_list = []
	# result_of_search = []
	
	# print (table_dict)
	
	# if 'SEARCHVALUE' in table_dict:
		# for k,v in table_dict.items():
			# if k == 'SEARCHVALUE':
				# for i in range(element_length):
					# for j in range(element_length):
						# if str(v[i]) == str(v[j]):
							# result_of_search += 
				
	
	try:
		for k,v in table_dict.items():
			j = 0
			if str(k) in input_string:
				print(str(k))
				for i in range(element_length):
					if j in index_of_final_list:
						regex = re.escape(str(k)) + '=\-?\"?[A-Za-z0-9]*\"?'
						restring = re.sub(regex, re.escape(str(k)) + '=' + str(v[i]), final_script_list[j])
						print(restring)
						# input_string = str(restring)
						final_script_list[j] = restring

					else:
						regex = re.escape(str(k)) + '=\-?\"?[A-Za-z0-9]*\"?'
						restring = re.sub(regex, re.escape(str(k)) + '=' + str(v[i]), input_string)
						# input_string = str(restring)
						final_script_list.insert(j, restring)
						index_of_final_list += [j]
					j+=1
						#print(str(i) + '\t' + str(regex) + str(v[i]) + ' ' + restring)

	except TypeError:
		pass
	#final_script_list.reverse()
	count_file = 0
	os.chdir('C:/Users/IABogdanov/Documents/py/gen_scripts/')
	files_in_dir = os.listdir()
	while str('S_' + str(count_file) + '_' + str(get_output_name())) in files_in_dir:
		count_file += 1
	with open('S_' + str(count_file) + '_' + str(get_output_name()), 'w') as f:
		for i in range(len(final_script_list)):
			print(final_script_list[i])

			f.write(final_script_list[i] + '\n')
		print('OK')
	f.close
def get_output_name() -> str:#прописать исключение вместо if/else
	#c_date = 'default ' + str(datetime.date.today()) + '.txt'
	c_date = str(datetime.date.today()) + '.txt'
		
	return c_date
	
def get_string() -> str:
	print('\nКоманда: ')
	input_string = input()
	
	return input_string






#rstring = re.sub('[,]\w*=\w*\d*[,;]+', ',TRXID=3413', tstring)
#---------------main----------------
os.chdir('C:/Users/IABogdanov/Documents/py/files/')
generate_script()

