# -*- coding: UTF-8 -*-
#!/usr/bin/env python

import win32com.client

class WordReport:
	""" Wrapper around Word 8 documents to make them easy to build.
		Has variables for the Applications, Document and Selection;
		most methods add things at the end of the document.
		Taken from 'Python Programming on Win32'
	"""
	def __init__(self, templatefile=None):
		self.wordApp = win32com.client.Dispatch('Word.Application')
		if templatefile == None:
			self.wordDoc = self.wordApp.Documents.Add()
		else:
			self.wordDoc = self.wordApp.Documents.Add(Template=templatefile)

		#set up the selection
		self.wordDoc.Range(0,0).Select()
		self.wordSel = self.wordApp.Selection

	def show(self):
		# convenience when debugging
		self.wordApp.ActiveWindow.View.Type = 3     # wdPageView
		self.wordApp.Visible = 1

	def getStyleList(self):
		# returns a dictionary of the styles in a document
		self.styles = []
		stylecount = self.wordDoc.Styles.Count
		for i in range(1, stylecount + 1):
			styleObject = self.wordDoc.Styles(i)
			self.styles.append(styleObject.NameLocal)

	def saveAs(self, filename):
		self.wordDoc.SaveAs(filename)

	def printout(self):
		self.wordDoc.PrintOut()

	def selectEnd(self):
		# ensures insertion point is at the end of the document
		self.wordSel.Collapse(0)
		# 0 is the constant wdCollapseEnd; don't want to depend
		# on makepy support.

	def selectBegin(self):
		# ensures insertion point is at the start of selection
		self.wordSel.Collapse()
		# default of Collapse() is start

	def addText(self, text, bookmark=None):
		if bookmark:
			self.wordDoc.Bookmarks(bookmark).Select()
			self.wordApp.Selection.TypeText(text)
		else:
			self.wordSel.InsertAfter(text)
		self.selectEnd()

	def addStyledPara(self, text, stylename):
		if text[-1] <> '\n':
			text = text + '\n'
		self.wordSel.InsertAfter(text)
		self.wordSel.Style = stylename
		self.selectEnd()