import socket
from collections import OrderedDict
import time
from math import *
import numpy as np
import numexpr as ne
import win32com.client



def SeparatorIndex(message,element,count):
	i=0
	j=0
	while i<len(message):
		if message[i]==element:
			j+=1
		if j==count:
			break
		i+=1
	return i

def CleanerStrToList(stringData):
	deletedSymbols=['[',']','\t','\n','\'',' ']
	i=0
	while i<len(deletedSymbols):
		stringData=stringData.replace(deletedSymbols[i], '')
		i+=1
	stringData=stringData.split(',')
	i=0
	while i<len(stringData):
		stringData[i]=float(stringData[i])
		i+=1
	return stringData

def ReadAnswer(listSocket,listDevices):
        
    message=''
    data=''
    while True:
        try:
            message = listSocket[listDevices].recv(100).decode()
        except socket.timeout:
            print("\nПрибор не отвечает\n")
            break
        except AttributeError:
            print("\nПрибор не отвечает\n")
            break
        last=len(message)
        data+=message  
        if message[last-1] == "\n":
            break
    return data

def WriteCommand(scenary,choice,sockets,devices):
	buffer = []
	start=SeparatorIndex(scenary, '=', 1)+1
	end=SeparatorIndex(scenary,'\n', 1)
	tmp = ''
	
	while start<end:
		tmp+=scenary[start]
		start+=1
	tmp = tmp.replace('\n',',')
	tmp=tmp.split(',')
	if choice=='График=':
		buffer.append('Координаты по X;Координаты по Y;\n')
	
	j=-1
	while j < len(tmp):
		if tmp[j] == '':
			del tmp[j]
		j+=1
	j=0
	res = ''
	while j < len(tmp):
		print('\n	Отправка команды: \"*OPC?;\"' + '\n')
		sockets[devices].send(('*OPC?;'+'\n').encode())
		res = str(ReadAnswer(sockets,devices))
		print('	Ответ от команды *OPC?; : \"' + str(res) + '\"\n')
		if res!='0' or res!='0\n' or res!=False or res!='+0\n' or res!='+0':
			print('\n	Отправка команды: \"' + str(tmp[j]) + '\"\n')
			sockets[devices].send((tmp[j]+'\n').encode())
			if tmp[j][len(tmp[j])-2]=='?':
				if choice=='Отправить':
					buffer.append('	')
				buffer.append(ReadAnswer(sockets,devices))
				print('	Ответ от команды: \"' + str(buffer) + '\"\n')
				if choice=='Отправить':
					buffer.append('\n\n\n')
		j+=1
	
	return buffer


def ConvertMeas(strr):
	unit_meas=1
	if strr[SeparatorIndex(strr,' ',1)+1:]=='КГц':
		unit_meas=(10**3)
	if strr[SeparatorIndex(strr,' ',1)+1:]=='МГц':
		unit_meas=(10**6)
	if strr[SeparatorIndex(strr,' ',1)+1:]=='ГГц':
		unit_meas=(10**9)
				
	strr=float(strr[:SeparatorIndex(strr,' ',1)])*unit_meas
	return strr
	
def Main():
	file = open('Конфиг.txt', 'r')
	config = file.readlines()
	
	
	if len(config)==0:
		print('\n	файл \'Конфиг.txt\'  пустой\n')
		return
	file.close()
	i=0
	while i < len(config):
		config[i] = config[i].replace('\n','')
		i+=1
	
	start=SeparatorIndex(config[0],'=',1)
	
	if 'Таймаут=' != config[0][:start+1]:
		print('\n	В файле \'Конфиг.txt\' в первой строке не найдена команда - \'Таймаут=\' \n')
		return
	tmp=0
	tmp=config[0][start+1:]
	
	tmp=tmp.replace(' ','')
	tmp = int(tmp)

	if tmp==0 or type(tmp)!=int or tmp=='':
		print('\n	В файле \'Конфиг.txt\' в команде - \'Таймаут=\' \n')
		print('	неверно указано значение \n')
		return
	
	
	
	start=SeparatorIndex(config[1],'=',1)
	
	if 'Количество приборов=' !=config[1][:start+1]:
		print('\n	В файле \'Конфиг.txt\' во второй строке не найдена команда - \'Количество приборов=\' \n')
		return
	tmp2=0
	tmp2=config[1][start+1:]
	
	tmp2=tmp2.replace(' ','')
	tmp2 = int(tmp2)+1
	if tmp2==0 or type(tmp2)!=int or tmp2=='':
		print('\n	В файле \'Конфиг.txt\' в команде - \'Количество приборов\'')
		print('	неверно указано значение \n')
		return
	
	
	i=2
	sockets=[]

	devices=[]
	ip=[]
	port=[]
	while i < tmp2+1:
		
		if not '#' in config[i]:
			start=SeparatorIndex(config[i],'=',1)
			devices.append(config[i][:start])
		
			start=SeparatorIndex(config[i],'=',1)
			end=SeparatorIndex(config[i],',',1)
			ip.append(config[i][start+1:end])
		
			port.append(int(config[i][end+1:]))
		
			sockets.append(socket.socket(socket.AF_INET,socket.SOCK_STREAM))

		i+=1
	i=0
	while i < len(sockets):
	
		sockets[i].settimeout(tmp)
		i+=1
	i=0	
	print('\n	Измерительные приборы = '+str(devices))
	print('	IP-адреса измерительных приборов  = ' + str(ip))
	print('	Порты измерительных приборов  = '+ str(port) + '\n\n')
	
	i=0
	bv=True
	messagesErrors=[]
	messagesSuccess=[]

	while i < len(sockets):
		try:
			sockets[i].connect((str(ip[i]), int(port[i])))
			messagesSuccess.append(i)
		except:
			messagesErrors.append(i)
		i+=1
	i=0
	
	while i < len(messagesSuccess):
		print('	Успешное подключение к '+'\''+str(devices[messagesSuccess[i]])+'\'')
		i+=1
	i=0
	#print('\n\n')
	while i < len(messagesErrors):
		print('	Не удалось подключится к '+'\''+str(devices[messagesErrors[i]])+'\'')
		i+=1
	

	
	if len(messagesErrors)>0:
		return

	file=open('Сценарий.txt','r')
	scenary = file.read()
	if len(scenary)==0:
		print('\n	файл \'Сценарий.txt\'  пустой\n')
		return
	file.close()
	scenary = scenary.replace('\t','')
	scenary = scenary.split('\n')
	e=0
	while e<1000:
		i=0
		while i<len(scenary):
			if scenary[i]=='':
				del scenary[i]
				
			i+=1
		e+=1
		
	e=0
	i=0
	if len(scenary)==1:
		scenary.append('\n')

	file=open('Операции.txt','r')
	operations = file.read()
	
	if len(operations)==0:
		print('\n	файл \'Операции.txt\'  пустой\n')
		return
	file.close()
	operations = operations.replace('\t','')
	operations= '\n'.join(OrderedDict((w,w) for w in operations.split('\n')).keys())
	operations=operations.split('\n')
	
	i=0
	while i<len(operations):
		if operations[i]=='':
			del operations[i]
		
		i+=1
	i=0
	j=0
	buffer_m = []
	operationIndexes = [] # Индексное расположение операций в файле сценария
	operationValues = [] # Инициализация операций в файле сценария
	namesOperations = [] # Вызов операций в файле сценария
	namesDevices = [] # Вызов операций для конкретного измерительного прибора в файле сценария
	while i<len(operations):
		j=0
		while j<len(scenary):
			if operations[i] in scenary[j] and not '#'  in scenary[j] and not '='  in scenary[j]:
				operationIndexes.append(j)
				operationValues.append(scenary[j])
			j+=1
		
		i+=1

	ind=0
	z=0
	tr=False
	j=0
	i=0
	while i<len(scenary): # 2.3.
		j=0
		while j<len(devices):
			if  devices[j] in scenary[i]  and not '#'  in scenary[i] and '='  in scenary[i]:

				if scenary[i][SeparatorIndex(scenary[i],'=',1)+1:] in operations:
					namesOperations.append(scenary[i][SeparatorIndex(scenary[i],'=',1)+1:])
					namesDevices.append(scenary[i][:SeparatorIndex(scenary[i],'=',1)])
					ind+=1

			j+=1
		i+=1
		
		
		
		
	j=0
	i=0
	while i<len(scenary):
		if 'Запуск программы' in scenary[i]:
			operationIndexes.append(i)
			operationValues.append(scenary[i])
		i+=1

	j=0
	i=0
	
	while i<len(operationIndexes):
		j=i+1
		while j<len(operationIndexes):
			if operationIndexes[i]>=operationIndexes[j]:
				trnp=operationValues[i]
				operationValues[i]=operationValues[j]
				operationValues[j]=trnp
				
				trnp=operationIndexes[i]
				operationIndexes[i]=operationIndexes[j]
				operationIndexes[j]=trnp
			j+=1
		i+=1
	
	i=0
	j=0
	if operationValues[len(operationValues)-1]!='Запуск программы': 
		print(' Не указан Запуск программы в файле сценария')
		return
	i=0
	j=0

	print('\n\n\n\n')

	counterOperations=0
	counterScenary=0
		
	VectorX=[]
	VectorY=[]
	tr=False
	while counterOperations<len(namesOperations):
		counterScenary=0
		h=0	
		endd=0
		slp=0.0
		while counterScenary<len(scenary):
			
			if tr==True:
				tr=False
				break
			if namesOperations[counterOperations]==scenary[counterScenary]:
				z=0
				while z<len(operationValues):
					
					if scenary[counterScenary]==operationValues[z]:
						h=operationIndexes[z]
						endd=operationIndexes[z+1]
						tr=True
						
						buffer_m.append('\n' + '	Измерительный прибор - \' '+ namesDevices[counterOperations] +'\'\n	Операция - \''+str(scenary[h])+'\'\n')
						print('\n' + '	Измерительный прибор - \''+ namesDevices[counterOperations] +'\'\n	Операция - \''+str(scenary[h])+'\'\n')
						
						break
						
					z+=1
			counterScenary+=1
		VALUES = []
		EXPRESSIONS = []
			
		
		while h<endd:
			
			if not '#' in scenary[h] and 'Задержка='  in scenary[h]: # в секундах
				
				namess=float(scenary[h][SeparatorIndex(scenary[h],'=',1)+1:])
				
				if type(namess)!=float or namess=='':
					print('\n	В файле \'Сценарий.txt\' в команде - \'Задержка\'')
					print('	неверно указано значение \n')
					return
				slp=namess
				del namess
				
				
			if not '#' in scenary[h] and 'Открыть таблицу='  in scenary[h]: 
				
				excel_path=scenary[h][SeparatorIndex(scenary[h],'=',1)+1:]
				excel_path=excel_path.replace('\n','')

				Excel = win32com.client.Dispatch("Excel.Application")
				wb = Excel.Workbooks.Open(u''+excel_path)
				sheet = wb.ActiveSheet
		
			if not '#' in scenary[h] and 'Сортировка('  in scenary[h]:

				namess=str(expression[SeparatorIndex(expression,'(',1)+1:SeparatorIndex(expression,'(',2)-1])

						 
				if namess=='VectorX':
					VectorX=sorted(VectorX)
							
				if namess=='VectorY':	
					VectorY=sorted(VectorY)
				
				del namess


			if  not '#' in scenary[h] and 'Формула('  in scenary[h] :
			
			
				value = scenary[h][SeparatorIndex(scenary[h],'(',1)+1:SeparatorIndex(scenary[h],';',1)]
				print('\n Переменные формулы : ' + str(value))
						
				value = value.replace(' ','')
				value=value.split(',')
						
				expression=scenary[h][SeparatorIndex(scenary[h],';',1)+1 : len(scenary[h])-1]
				print('\n Выражение : ' + str(expression))
						
				expression = expression.replace(' ','')
				
				expression = expression.replace('^','**')	
				exc = expression[0:SeparatorIndex(expression,'=',1)+1]
						
				if not 'VectorXY' in expression:
						
					EXPRESSIONS.append(expression[0:SeparatorIndex(expression,'=',1)])
					b=0
					expression = expression[SeparatorIndex(expression,'=',1)+1:]
						
				if 'VectorXY' in expression:
					z=0
					VectorX=[]
					VectorY=[]
					xy=[]
					while z<len(buffer_m):
						if  not '\t' in buffer_m[z] and not '\n\n\n' in buffer_m[z] and buffer_m[z]!='\n':
							xy.append(buffer_m[z])
						z+=1
							
					z=0
					zz=0
					while z<len(operations):
						zz=0
						while zz<len(xy):
							if operations[z] in xy[zz]:
								del xy[zz]
								
							zz+=1
						z+=1

					VectorX=CleanerStrToList(xy[0][:len(xy[0])-1])
					VectorY=CleanerStrToList(xy[1][:len(xy[1])-1])
							
					print('VectorX\n\n   ' + str(VectorX)+'\n\nVectorX\n\n' )
					print('VectorY\n\n   ' + str(VectorY)+'\n\nVectorY' )

					
				if not 'VectorXY' in expression:
					gl_expr=''
					while b<len(value):
						expression = expression.replace(value[b][0:SeparatorIndex(value[b],'=',1)] , str(value[b][SeparatorIndex(value[b],'=',1)+1:]))
							
						b+=1
					b=0
						
					if len(VALUES)>0 :
						while b<len(VALUES):
							expression = expression.replace(EXPRESSIONS[b],str(VALUES[b]))
							b+=1
							
					if 'от' in expression and 'до' in expression and len(VectorX)>0 and 'по' in expression:

						fnnn=str(expression[SeparatorIndex(expression,'т',1)+1 : SeparatorIndex(expression,'д',1)])
								
						fvvv=str(expression[SeparatorIndex(expression,'о',2)+1 : SeparatorIndex(expression,'п',1)])
									
						veccc=str(expression[SeparatorIndex(expression,'о',3)+1: SeparatorIndex(expression,'V',2)])
								
						veccc2=str(expression[SeparatorIndex(expression,'V',2):])
								
						expression = '('+ fnnn +'<'+veccc+')&('+veccc+'<'+fvvv+')'
						gl_expr=expression

						ans=ne.evaluate(expression)

						b=0
						xxxx=[]
						yyyy=[]
								
						while b<len(ans):
							if ans[b]==True:
								if veccc2 == 'VectorY':
									yyyy.append(VectorY[b])
												
								if veccc2 == 'VectorX':
									xxxx.append(VectorX[b])
							b+=1
						if veccc2 == 'VectorX':
								VectorX=xxxx
									
						if veccc2 == 'VectorY':	
								VectorY=yyyy
									

						if veccc2=='VectorX':
							print('\n Результаты вычислений : ' + str(exc) + str(VectorX) +'\n')
							buffer_m.append('\n	'+str(exc) + str(VectorX)+'\n')
								
						if veccc2=='VectorY':
							print('\n Результаты вычислений : ' + str(exc) + str(VectorY) +'\n')
							buffer_m.append('\n	'+str(exc) + str(VectorY)+'\n')
						b=0
								
					b=0	

					print('\n Вычисление формулы : '+ str(exc) + str(expression))
						
					try:
						
						
						VALUES.append(ne.evaluate(expression))
						if not '&' in expression:
							print('\n Результаты вычислений : ' + str(exc) + str(ne.evaluate(expression)) +'\n')
							buffer_m.append('\n	'+str(exc) + str(ne.evaluate(expression))+'\n')		
					except:
						print('\nОшибка в формуле, возможная причина - имя переменной было заданно неверно или отсутствует.\n')
						return


			if  not '#' in scenary[h] and 'Измерение с шагом='  in scenary[h]  and scenary[h+1]=='[':
				
				str_write=''
				ch_freq2=h
				chj=h
				str_write_size1=0
				str_write_size2=0
				while(chj<endd):
					if scenary[chj]=='[':
						str_write_size1=chj
						break
		
					chj+=1
					
				ch_freq2=h
				chj=h				
				while(chj<endd):
					if scenary[chj]==']':
						str_write_size2=chj
						break

					chj+=1
				ch_freq2=str_write_size1+1
			
				size_meas=float(scenary[h][SeparatorIndex(scenary[h],'=',1)+1 : ])
				
				ch_freq2=str_write_size1+1
				step_freqs=[]
				start_freqs=[]
				while ch_freq2<str_write_size2:
					
					if not '#' in scenary[ch_freq2] and 'Отправить='  in scenary[ch_freq2]:
						tmp_prb=scenary[ch_freq2][ SeparatorIndex(scenary[ch_freq2],',',2)+1 : SeparatorIndex(scenary[ch_freq2],',',3) ] 
					
						if '...' in tmp_prb:
							tmp_start_freq=tmp_prb[ : tmp_prb.find('...')]
							tmp_start_freq=ConvertMeas(tmp_start_freq)
						
							tmp_end_freq=tmp_prb[ tmp_prb.find('...')+3 : ]
							tmp_end_freq=ConvertMeas(tmp_end_freq)
						
							start_freqs.append(tmp_start_freq)
							step_freqs.append(((tmp_end_freq-tmp_start_freq)/size_meas))
					
						if tmp_prb=='0':
							start_freqs.append(tmp_prb)
							step_freqs.append(tmp_prb)
						
						if not '...' in tmp_prb and tmp_prb!='0' :	
					
							tmp_start_freq=tmp_prb
							tmp_start_freq=ConvertMeas(tmp_start_freq)
						
							start_freqs.append(tmp_start_freq)
							step_freqs.append(0)
					
					ch_freq2+=1
				
				ch_freq=0
				flg=False
				local_step=0
				add_meas=0
				if size_meas>1:
					add_meas=1
				if str_write_size1>0 and size_meas>0 :
					
					while(ch_freq<size_meas+add_meas):
						ch_freq3=str_write_size1+1
						local_step=0
						tmp_prb=scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],'=',1)+1 : SeparatorIndex(scenary[ch_freq3],',',1)]
						print("\n\n	Выполнение команды - \'Измерение с шагом\', для прибора\'" + str(tmp_prb) +" Гц, Точка измерения="+str(ch_freq)+'\n\n')
						vib=''
						while(ch_freq3<str_write_size2):
							buffer=[]
							
							if not '#' in scenary[ch_freq3] and 'Отправить='  in scenary[ch_freq3]:
								tmp_prb=scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],'=',1)+1 : SeparatorIndex(scenary[ch_freq3],',',1)]
								fvn=0
							
								pribor=0
								while fvn<len(devices):
					
									if devices[fvn]==tmp_prb:
										pribor=fvn
										
										break
							
									fvn+=1
								tmp_prb=''
								if ' шаг ' in scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',3)+1 : ]:
								
									str_write=scenary[ch_freq3]
									str_write=str_write[ : SeparatorIndex(str_write,'=',1)+1 ] + str_write[SeparatorIndex(str_write,',',3)+1 : ].replace(' шаг ', ' '+str(start_freqs[local_step])+' ')
									vib='Отправить'
									
							
									
								if not ' шаг ' in scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',3)+1 : ]:
									str_write=scenary[ch_freq3]
									str_write=str_write[ : SeparatorIndex(str_write,'=',1)+1 ] + str_write[SeparatorIndex( str_write,',',3)+1 : ]								
									
								if scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]== 'Без таблицы':
									vib='Отправить'
										
								if 'Заголовок' in scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]:
									vib='none'
									
							
								buffer=WriteCommand(str_write,vib,sockets,pribor)
								if len(buffer)>0:
									buffer_m+=buffer
								e=0
								while e<len(buffer_m):
									buffer_m[e] = buffer_m[e].replace('[','')
									buffer_m[e] = buffer_m[e].replace(']','')
									e+=1
								
								
								
								if ' шаг ' in scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',3)+1 : ] and 'Заголовок' in scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]:
										
									
									
									row_col=scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]
									row_col=int(row_col[SeparatorIndex(row_col,'=',1)+1 : SeparatorIndex(row_col,';',1)])
									
									
									if flg==False:
										header_excel=scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]
										header_excel=header_excel[SeparatorIndex(header_excel,';',1)+1 : ]+
										sheet.Cells(2,row_col).value = header_excel

									sheet.Cells(3+ch_freq,row_col).value = start_freqs[local_step]
									vib='Отправить'
								
								
								
								
								if not ' шаг ' in scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',3)+1 : ] and 'Заголовок' in scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]:
									

									e=0
									buffer[e] = buffer[e].replace('[','')
									buffer[e] = buffer[e].replace(']','')
									buffer[e] = buffer[e].replace('\t','')
									buffer[e] = buffer[e].replace('\n','')
									

									row_col=scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]
									row_col=int(row_col[SeparatorIndex(row_col,'=',1)+1 : SeparatorIndex(row_col,';',1)])
									
									
									if flg==False:
										header_excel=scenary[ch_freq3][SeparatorIndex(scenary[ch_freq3],',',1)+1 : SeparatorIndex(scenary[ch_freq3],',',2)]
										header_excel=header_excel[SeparatorIndex(header_excel,';',1)+1 : ]+
										sheet.Cells(2,row_col).value = header_excel
									sheet.Cells(3+ch_freq,row_col).value = buffer
							
								print('\n\n  Задержка между командами='+ str(slp) + ' сек.')
								time.sleep(slp)	
							
								start_freqs[local_step]+=step_freqs[local_step]
								local_step+=1
							
							ch_freq3+=1
						
						flg=True	
						ch_freq+=1
						

			if not '#' in scenary[h] and 'Закрыть таблицу=Закрыть'  in scenary[h]: 
				
				wb.Save()

				#закрываем ее
				wb.Close()

				#закрываем COM объект
				Excel.Quit()
				#return
			
			
			
			if  not '#' in scenary[h] and 'Сохранение='  in scenary[h] :
				start=SeparatorIndex(scenary[h],'=',1)
				end=SeparatorIndex(scenary[h],',',1)
				fl=[]
				fl.append(scenary[h][start+1:end])
					
				if fl[0]=='Перезаписать':
					start=SeparatorIndex(scenary[h], ',', 1)
					end=SeparatorIndex(scenary[h],'\n', 1)
					file = open(scenary[h][start+1:end], 'w')
	
				
					if len(buffer_m)>0:
					
						file.writelines(buffer_m)
						file.close()
						buffer_m=[]
			
				if fl[0]=='Дописать':
					start=SeparatorIndex(scenary[h], ',', 1)
					end=SeparatorIndex(scenary[h],'\n', 1)
					file = open(scenary[h][start+1:end], 'a')
					
				
					if len(buffer_m)>0:
					
						file.writelines(buffer_m)
						file.close()
						buffer_m=[]
			
				if fl=='':
					print('\n	Ошибка: Пустой аргумент в вызове функции')

			h+=1
		print('\n\n  Задержка между операциями='+ str(slp) + ' сек.')
		time.sleep(slp)	
		print('\n\n\n\n')
	
		counterOperations+=1

Main()	
print('\n	Нажмите \'Enter\' для выхода из программы ')	
input()	
	
	
	
