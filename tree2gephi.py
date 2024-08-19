from tkinter import Tk 
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
import pandas as pd

def calc_levels(headers):
	# Calcula la cantidad de niveles en el excel de entrada 
	counter = 0
	for header in headers[1:]:
		if header == "Company Name" or "Level" in header:
			counter += 1
			continue
		break
	return counter

def create_nodes(rows, levels):
	# Crea una lista de nodos
	result = []

	for row in rows[6:]:
		values = [cell.value for cell in row]
		labels = values[1:levels+1]
		values = [values[0]] + values[levels+1:]
		label = next(s for s in labels if s)
		result_row = {
			"Id": values[0],
			"Label": label,
			"Type": values[2],
			"Country/Region": values[3],
			"Incorporated Country": values[4],
			"Incorporated Date": values[5],
			"PE Backed Status": values[6],
			"Industry": values[7],
			"Total Revenue": values[8],
			"Employees": values[9],
			"Market Cap": values[10],
			"Implied Rating": values[11],
			"Moody's Rating": values[12],
			"Fitch Rating": values[13],
		}
		result.append(result_row)
	return result

def create_relations(rows, levels):
	# Crea una lista de relaciones
	result = []
	parent = [row[6][0].value]
	previous_level = -1
	previous_id = -1
	for row in rows[7:]:
		values = [cell.value for cell in row]
		labels = values[1:levels+1]
		id = values[0]
		level, label = next((level,s) for level, s in enumerate(labels) if s)
		if level == 0:
			continue
		if level > previous_level:
			parent.append(previous_id)
		if level < previous_level:
			parent.pop()
		result.append({
						"Source": parent[-1], 
						"Target": id, 
						"Relationship Type": values[5], 
						"Ownership Percentage": values[18]
					})
		previous_level = level
		previous_id = id
	return result

def create_output(rows, levels):
	nodos = create_nodes(rows, levels)
	relaciones = create_relations(rows[:6], levels)
	return nodos, relaciones

if __name__ == '__main__':
	# Primero abre una ventana de dialogo para elegir un archivo
	Tk().withdraw() 
	filename = askopenfilename() 
	
	# Lectura del excel de entrada
	wb = load_workbook(filename=filename)
	ws = wb.active
	rows = list(ws.rows)
	headers = [cell.value for cell in list(rows[5])]
	levels = calc_levels(headers)

	# Manipulaciones del excel de entrada
	nodos, relaciones = create_output(rows, levels)

	# Escritura de los exceles de salida
	df_nodos = pd.DataFrame.from_dict(nodos)
	df_nodos.to_excel('nodos.xlsx')
	df_relaciones = pd.DataFrame.from_dict(relaciones)
	df_relaciones.to_excel('relaciones.xlsx')
	