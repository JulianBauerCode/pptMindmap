import pandas as pd
import collections

#####################################
# Definitions

class Interaction(object):
    # def __init__(self, nameProject):
        # self.nameProject = nameProject
        # self.labelTo = None
        # self.labelFrom = None
        # self.labelWith = None
        
    def __init__(self):
        self.labelTo = None
        self.labelFrom = None
        self.labelWith = None
        
    def addTo(self, label):
        self.labelTo = label

    def addFrom(self, label):
        self.labelFrom = label

    def addWith(self, label):
        self.labelWith = label
    
    def __repr__(self,):
        repr = ''
        for key, value in self.__dict__.items():
            if value:
                repr = repr + '{}:{}  '.format(key, eval('self.'+key))
        return repr

#####################################
# Parameters

blank = 'x'

#####################################
# Read data

supportsExcel = pd.read_excel('supports.xlsx', index_col=0)
cooperationsExcel = pd.read_excel('cooperations.xlsx', index_col=0)

# Check consistency
assert supportsExcel.columns.tolist() == cooperationsExcel.columns.tolist()
assert supportsExcel.index.values.tolist() == supportsExcel.columns.tolist()
assert cooperationsExcel.index.values.tolist() == cooperationsExcel.columns.tolist()

projects = supportsExcel.columns.tolist()

#####################################
# Translate input data into comfortable format

interactions = {}
for project in projects:
    interactions[project] = collections.defaultdict(Interaction)

import collections

# supporting interactions
for rowTo in projects:
    for colFrom in projects:
        label = supportsExcel.loc[rowTo][colFrom]
        if label != blank:
            interactions[rowTo][colFrom].addFrom(label)
            interactions[colFrom][rowTo].addTo(label)

# supporting cooperations
for rowTo in projects:
    for colFrom in projects:
        label = cooperationsExcel.loc[rowTo][colFrom]
        if label != blank:
            interactions[rowTo][colFrom].addWith(label)
            interactions[colFrom][rowTo].addWith(label)

del label, rowTo, colFrom, project, blank, cooperationsExcel, supportsExcel

#####################################
# Powerpoint

import win32com.client
import sys
import os

fileName = 'graph.pptx'
filePath = os.path.join(os.getcwd(), fileName)

app = win32com.client.Dispatch("PowerPoint.Application")
#app.Visible = False
p = app.Presentations.Open(filePath, WithWindow=False)


p.Save()
#p.SaveAs(filePath+'test')
p.Close()
app.Quit()