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
import numpy as np


fileName = 'graph.pptx'
filePath = os.path.join(os.getcwd(), fileName)


slideWidth = 25.4 # cm
slideHeight = 19.05 # cm


app = win32com.client.Dispatch("PowerPoint.Application")
#app.Visible = False
p = app.Presentations.Open(
                        filePath,
                        WithWindow=False,
                        ReadOnly=False,
                        )
s = p.slides[0]

class ConnectionGraph(object):

    pointsPerCm = 28.35

    def __init__(self, slide, centralLabel, centralCoords, interactions, sizesDict_centralKids_WidthHeight):
        self.slide = slide
        self.centralLabel = centralLabel
        self.centralCoords = [item * self.pointsPerCm for item in centralCoords]
        self.interactions = interactions
        self.sizes = self.dictValuesFromCentimeterToPoint(
                        d=sizesDict_centralKids_WidthHeight,
                        pointsPerCm=self.pointsPerCm,
                        )

    def dictValuesFromCentimeterToPoint(self, d, pointsPerCm):
        '''Multiply all values of dict which are numbers by pointsPerCm
        '''
        from numbers import Number
        from collections import Mapping
        for k, v in d.items():
            if isinstance(v, Mapping):
                self.dictValuesFromCentimeterToPoint(v, pointsPerCm)
            elif isinstance(v, Number):
                d[k] = v * pointsPerCm
        return d

    def xyByPolar(self, radius, angle):
        angle = np.deg2rad(angle)
        return (radius * np.cos(angle), radius * np.sin(angle))

    def drawcentral(self,):
        width = self.sizes['central']['width']
        height = self.sizes['central']['height']
        box = self.slide.Shapes.AddTextbox(
            Orientation=1,# msoTextOrientationHorizontal
            Left=self.centralCoords[0]-width/2.0,
            Top=self.centralCoords[1]-height/2.0,
            Width=width,
            Height=height,
            )
        box.TextFrame.TextRange.Text = self.centralLabel
        box.TextEffect.Alignment = 2    # centered

    def drawKids(self,):
        width = self.sizes['kids']['width']
        height = self.sizes['kids']['height']
        nbrKids = len(self.interactions)
        angleIncrement = 360.0 / float(nbrKids)
        radius = 2.0 * self.pointsPerCm

        # for kidKey, kidValue in self.interactions.items():
        kidKey, kidValue = list(self.interactions.items())[0]
        print(kidValue)
        angle = angleIncrement * 0
        coords = self.xyByPolar(radius, angle)
        box = self.slide.Shapes.AddTextbox(
            Orientation=1,# msoTextOrientationHorizontal
            Left=self.centralCoords[0] + coords[0]-width/2.0,
            Top=self.centralCoords[1] + coords[1]-height/2.0,
            Width=width,
            Height=height,
            )
        box.TextFrame.TextRange.Text = kidKey
        box.TextEffect.Alignment = 2    # centered

    def drawToLabels(self,):
        pass

    def drawFromLabels(self,):
        pass

    def drawWithLabels(self,):
        pass

g = ConnectionGraph(
        slide=s,
        centralLabel='S1',
        centralCoords=(slideWidth/2.0, slideHeight/2.0),
        interactions=interactions['S1'],
        sizesDict_centralKids_WidthHeight={
            'central':{
                'width':2,  # cm
                'height':1, # cm
                },
            'kids':{
                'width':2,  # cm
                'height':1, # cm
                },
            },
        )
g.drawcentral()
g.drawKids()

# p.Save()
p.SaveAs(filePath+'test')
p.Close()
app.Quit()