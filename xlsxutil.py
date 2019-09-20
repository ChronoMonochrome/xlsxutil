import tempfile, shutil, os
import io
import zipfile
import uuid

from copy import deepcopy

from lxml import etree
import xml.etree.ElementTree as ET

from utils import create_temporary_copy, n, UpdateableZipFile

class Cell:
	def __init__(self, cell_el, sheet):
		self._sheet = sheet
		self._cell = cell_el
		
	@property
	def value(self):
		return self._cell.find(n("v")).text
		
	@value.setter
	def value(self, v):
		self._sheet._set_dirty(True)
		if type(v) != str:
			v = str(v)

		self._cell.find(n("v")).text = v
		
class Cells:
	def __init__(self, cells, sheet):
		self._sheet = sheet
		self._cells = cells
		 
	def __getitem__(self, idx):
		return Cell(self._cells[idx], self._sheet)

class Row:
	def __init__(self, row_el, sheet):
		self._sheet = sheet
		self._row = row_el
		self._cells = self._row.findall(n("c"))
		
	@property
	def cells(self):
		"""for cell in self._cells:
			   yield Cell(cell)"""
		return Cells(self._cells, self._sheet)
			   
class Rows:
	def __init__(self, rows, sheet):
		self._rows = rows
		self._sheet = sheet
		 
	def __getitem__(self, idx):
		return Row(self._rows[idx], self._sheet)

class Sheet:
	_initialized = False
	_dirty = False

	def __init__(self, workbook, sheet_name):
		self._workbook = workbook
		self.name = sheet_name
		
	def _set_dirty(self, dirty):
		self._dirty = dirty
		
	def lazy_init(self):
		if not self._initialized:
			self._sheet_id =  self._workbook._worksheets[self.name]
			self._sheet_path = 'xl/worksheets/sheet%s.xml' % self._sheet_id
			self._tree = ET.parse(self._workbook.fh.open(self._sheet_path))
			self._root = self._tree.getroot()
			self._rows = self._root.findall(n("sheetData")+"/"+n("row"))
			self._initialized = True
			
	def save(self, file_path):
		if not self._dirty:
			return
			
		with UpdateableZipFile(file_path, mode = "a") as fd:
			fd.remove_file(self._sheet_path)
		 
		xml_data = ET.tostring(self._root, encoding='utf8', method='xml')
		
		# A dirty hack to get excel working
		xml_data = xml_data.replace(b'''<ns0:worksheet xmlns:ns0="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:ns1="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ns2="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ns1:Ignorable="x14ac">''',
									b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'''
									).replace(b'''<?xml version='1.0' encoding='utf8'?>''',
											  b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'''
									).replace(b"ns0:", b"").replace(b"ns2:", b"x14ac:").replace(b"\n", b"\r\n")

		with UpdateableZipFile(file_path, mode = "a") as fd:
			fd.writestr(self._sheet_path, xml_data)
		
	@property
	def rows(self):
		"""for row in self._rows:
			yield Row(row)"""
		self.lazy_init()
		return Rows(self._rows, self)
		
class Workbook:
	def __init__(self, file_path):
		self.ns = {
			'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
		}
		self._tempfile_path = create_temporary_copy(file_path)
		self.fh = UpdateableZipFile(self._tempfile_path, mode = "a")
		#self.shared = self.load_shared()
		self._worksheets = self.load_workbook()
		self.worksheets = dict()
		
		for sheet_name, _ in self._worksheets.items():
			self.worksheets[sheet_name] = Sheet(self, sheet_name)

	def load_workbook(self):
		# Load workbook
		name = 'xl/workbook.xml'
		root = etree.parse(self.fh.open(name))
		res = {}
		for el in etree.XPath("//ns:sheet", namespaces=self.ns)(root):
			res[el.attrib['name']] = el.attrib['sheetId']
		return res

	def load_shared(self):
		# Load shared strings
		name = 'xl/sharedStrings.xml'
		root = etree.parse(self.fh.open(name))
		res = etree.XPath("/ns:sst/ns:si/ns:t", namespaces=self.ns)(root)
		return {
			str(pos): el.text
			for pos, el in enumerate(res)
		}
		
	def save(self, file_path):
		shutil.copy2(self._tempfile_path, file_path)

		for sheet_name in self.worksheets:
			self.worksheets[sheet_name].save(file_path)