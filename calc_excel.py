import win32com.client
from math import *
import numpy as np
import numexpr as ne
from numexpr import *
from numpy import *
def test(message,el,count):
	i=0
	j=0

	while i<len(message):
		if message[i]==el:
			j+=1
		if j==count:
			break
		i+=1
	return i

def row_excel(sheet):
	content_excel = [r[0].value for r in sheet.Range('A15:DZ15')]
	count=0
	row_content_excel=[]
	while count<len(content_excel):
			
		if content_excel[count]!=None and type(content_excel[count])==str:
			row_content_excel.append(content_excel[count])
			
		count+=1
	return row_content_excel
	

def sortt1D(vector1D):
	count=0
	count2=0
	while count<len(vector1D):
		count2=count+1
		while count2<len(vector1D):
			if vector1D[count]>=vector1D[count2]:
			
				trnp=vector1D[count]
				vector1D[count]=vector1D[count2]
				vector1D[count2]=trnp
			count2+=1
		count+=1
	return vector1D

def sortt2D(vector2D):
	
	count=0
	count2=0
	while count < len(vector2D):
	
		vector2D[count]=sortt1D(vector2D[count])
		count+=1
	return vector2D

	


def alphabet_index_excel(sheet):
	
	index_content_excel=[]
	count=0
	content_excel = [r[0].value for r in sheet.Range('A15:DZ15')]
	while count<len(content_excel):
		if content_excel[count]!=None:
			index_content_excel.append(count)
			
		
		count+=1

	#print(index_content_excel)
	
	return index_content_excel
	

	
def alphabet_excel():
	alphabet=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ','DA','DB','DC','DD','DE','DF','DG','DH','DI','DJ','DK','DL','DM','DN','DO','DP','DQ','DR','DS','DT','DU','DV','DW','DX','DY','DZ']
	return alphabet	
	

def count_headers_columns_excel(sheet):
	alphabet=alphabet_excel()
	count=0
	headers_columns_excel=[]
	#content_excel = [r[0].value for r in sheet.Range('A15:DZ15')]
	#content_excel = [r[0].value for r in sheet.Range('A15:DZ15')]
	#print(content_excel)
	#return 
	#index_content_excel=alphabet_index_excel(content_excel)
	index_content_excel=alphabet_index_excel(sheet)
	
	for count in range(len(index_content_excel)):
		headers_columns_excel.append([0]*0)
	
	count=0
	count2=0
	while count<len(index_content_excel):
		count2=0
		content_excel = [r[0].value for r in sheet.Range(str(alphabet[index_content_excel[count]])+'16:'+str(alphabet[index_content_excel[count]])+'500')]
		while count2<len(content_excel):
			
			if content_excel[count2]!=None:
				headers_columns_excel[count].append(float(content_excel[count2]))
			
			count2+=1
		
		count+=1
	count=0
	count2=0

	return headers_columns_excel


def testt(mas):
	return len(mas)

def main():	
	file=open('Расчет.txt','r')
	sluj = file.read()
	if len(sluj)==0:
		print('\n	файл \'Расчет.txt\'  пустой\n')
		return
	file.close()
	sluj = sluj.replace('\t','')

	sluj=sluj.split('\n')
	e=0
	while e<1000:
		i=0
		while i<len(sluj):
			if sluj[i]=='':
				del sluj[i]
				
			i+=1
		e+=1
		
	e=0
	i=0
	if len(sluj)==1:
		sluj.append('\n')
	
	
	buffer_main=[]
	
	while i<len(sluj):
	
		if not '#' in sluj[i] and 'Открыть таблицу='  in sluj[i]: 
				
			excel_path=sluj[i][test(sluj[i],'=',1)+1:] 
			excel_path=excel_path.replace('\n','')
			#print("\n\n\n	"+excel_path[len(excel_path)-1]+"\n\n\n")
			#return
			Excel = win32com.client.Dispatch("Excel.Application")
			wb = Excel.Workbooks.Open(u''+excel_path)
			sheet = wb.ActiveSheet	

			#excel_path='D:\\Python\\Stend_test\\2 controls\\python_excel\\Расчетный модуль\\xl.xlsx'
			#Excel = win32com.client.Dispatch("Excel.Application")
			#wb = Excel.Workbooks.Open(u''+excel_path)
			#sheet = wb.ActiveSheet
			
				
			

			
		if  not '#' in sluj[i] and 'Формула('  in sluj[i] :
			
			expression=sluj[i][test(sluj[i],'(',1)+1:len(sluj[i])-1]
			print('\n Вычисляемое выражение : ' + str(expression))
			expression=expression.replace(' ','')
			expression=expression.replace('.','')
			count_headers_columns=count_headers_columns_excel(sheet)
			row_content_excel=row_excel(sheet)
			dictt={}
			count=0
			while count<len(row_content_excel):
				row_content_excel[count]=row_content_excel[count].replace(' ','')
				row_content_excel[count]=row_content_excel[count].replace('.','')
				dictt.update({row_content_excel[count]:count_headers_columns[count]})
				count+=1
			#print(dictt.get())
			#return
			#if type(ne.evaluate(expression,dictt))!=float or type(ne.evaluate(expression,dictt))!=int or type(ne.evaluate(expression,dictt))!=bool:
			
			#tss=[1,2,3]
			#dictt.update({'tss':testt(tss)})
			#dictt.update({'testt_'+str(dictt.keys()):testt(dictt.values())})
			'''
			if 'len(' in expression and  ')' in expression:
				count=0
				tmp_str=expression[test(expression,'(',1)+1 : test(expression,')',1) ]
				#expression=expression.replace('(','_')
				#expression=expression.replace(')','_')
				
				expression=list(expression)
				expression[test(expression,'(',1)]='_'
				expression[test(expression,')',1)]='_'
				
				expression=''.join(expression)
			
				while count<len(row_content_excel):
					if row_content_excel[count]==tmp_str:
						dictt.update({'len_'+str(tmp_str)+ '_' : len(count_headers_columns[count])})
						#break
					#if count==len(row_content_excel)-1 and row_content_excel[count]!=tmp_str:
					#	print('\n\n	Введено неверное имя')
					#	return
					count+=1
				tmp_str=''
				#print('entire in condition')
				#print(expression)
				#return
			'''
			
			count=0
			user_func=['размер','сортировка']
			while count<len(user_func):
				if user_func[count] in expression and  ')' in expression:
					count2=0
					tmp_str=expression[test(expression,'(',1)+1 : test(expression,')',1) ]
					#expression=expression.replace('(','_')
					#expression=expression.replace(')','_')
				
					expression=list(expression)
					expression[test(expression,'(',1)]='_'
					expression[test(expression,')',1)]='_'
				
					expression=''.join(expression)
			
					while count2<len(row_content_excel):
						if row_content_excel[count2]==tmp_str:
							if user_func[count]=='размер':
								dictt.update({user_func[count]+'_'+str(tmp_str)+ '_' : len(count_headers_columns[count2])})
							if user_func[count]=='сортировка':
								dictt.update({user_func[count]+'_'+str(tmp_str)+ '_' : sorted(count_headers_columns[count2])})
							#break
						#if count==len(row_content_excel)-1 and row_content_excel[count]!=tmp_str:
						#	print('\n\n	Введено неверное имя')
						#	return
						count2+=1
				count+=1
				#print('entire in condition')
				#print(expression)
				#return
			
			
			
			
			
			#dictt.update({'testt_'+str(dictt.keys()):testt(dictt.values())})
			count=0
			
			tmp=[]
				
			
			#tmp=list(ne.evaluate(expression,dictt))
			#tmp=list(ne.evaluate(expression,dictt))
			try:
				#print(len(list(ne.evaluate(expression,dictt))))
				tmp=list(ne.evaluate(expression,dictt))
			except TypeError:
				tmp.append(ne.evaluate(expression,dictt))
				
					
				print('\n\n  TypeError: iteration over a 0-d array')
				
				
			if '<' in expression or '>' in expression:
				
				count=0
				index_row_content_excel=0
				while count<len(row_content_excel):
					
					if row_content_excel[count] in expression and len(count_headers_columns[count])>1:
						
						index_row_content_excel=count
						break
					count+=1
				
				count=0
				
				
				#print(count_headers_columns[index_row_content_excel])
				tmp2=[]	
				while count<len(count_headers_columns[index_row_content_excel]):
					if tmp[count]==True:
						tmp2.append(count_headers_columns[index_row_content_excel][count])
					count+=1
				
				tmp=tmp2
							
				
					
			count=0
			while count<len(tmp):
				buffer_main.append(float(tmp[count]))
				count+=1
			
			#if type(ne.evaluate(expression,dictt))!=list or type(ne.evaluate(expression,dictt))!=turple or type(ne.evaluate(expression,dictt))!=dict:
				#buffer_main.append(float(ne.evaluate(expression,dictt)))
				#print(buffer_main)
			wb.Save()
			
			#закрываем ее
			wb.Close()

			#закрываем COM объект
			Excel.Quit()
			
		if not '#' in sluj[i] and 'Сохранить в таблицу('  in sluj[i]: 
			excel_path=sluj[i][test(sluj[i],',',2)+1:len(sluj[i])-1]
			excel_path=excel_path.replace('\n','')
			#print("\n\n\n	"+excel_path[len(excel_path)-1]+"\n\n\n")
			#return
			Excel = win32com.client.Dispatch("Excel.Application")
			wb = Excel.Workbooks.Open(u''+excel_path)
			sheet = wb.ActiveSheet	

		
			if 'Заголовок' in sluj[i][test(sluj[i],'(',1)+1 : test(sluj[i],'=',1) ]:
			
			
				#row_col=int(sluj[i][test(sluj[i],'=',1)+1 : test(sluj[i],',',1) ])
				#row_col=int(row_col.replace(' ',''))
				row_col=sluj[i][test(sluj[i],'=',1)+1 : test(sluj[i],',',1) ]
				alphabet=alphabet_excel()
				count=0
				while count<len(alphabet):
					alphabet[count]=alphabet[count]+'15'
					count+=1
				
				
				count=0
				while count<len(alphabet):
					if alphabet[count]==row_col:
						row_col=count+1
						break
					count+=1	
				#print(row_col)
				#return
			
			#if sluj[i][test(sluj[i],',',1)+1 : test(sluj[i],'=',2) ]=='Название':
			if 'Название заголовка' in sluj[i][test(sluj[i],',',1)+1 : test(sluj[i],'=',2) ]:	
	
				header_excel=sluj[i][test(sluj[i],'=',2)+1 : test(sluj[i],',',2) ]
				
				
				sheet.Cells(15,row_col).value = header_excel
				
				#if type(ne.evaluate(expression,dictt))!=float or type(ne.evaluate(expression,dictt))!=int or type(ne.evaluate(expression,dictt))!=bool:
				
				count=0
				while count<len(buffer_main):
					sheet.Cells(16+count,row_col).value = buffer_main[count]
					count+=1	
				
				#if type(ne.evaluate(expression,dictt))!=list or type(ne.evaluate(expression,dictt))!=turple or type(ne.evaluate(expression,dictt))!=dict:
				
				#	sheet.Cells(3+1,row_col).value = buffer_main
				
			buffer_main=[]
			
			
			wb.Save()

			#закрываем ее
			wb.Close()

			#закрываем COM объект
			Excel.Quit()
			
		i+=1
		

main()






