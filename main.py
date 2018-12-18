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


class ConnectionGraph(object):

    pointsPerCm = 28.35

    def __init__(self, slide, centralLabel, centralCoords, interactions, sizesDict_centralKids_WidthHeight):
        self.slide = slide
        self.centralLabel = centralLabel
        self.centralCoords = self.dictValuesFromCentimeterToPoint(
                                    d=centralCoords,
                                    )
        self.interactions = interactions
        self.interactionKeys = interactions.keys()  # Do not rely on order of dicts
        self.sizes = self.dictValuesFromCentimeterToPoint(
                        d=sizesDict_centralKids_WidthHeight,
                        )
        self.centralShape = None
        self.kidShape = {}
        self.labelShape = {}

    def dictValuesFromCentimeterToPoint(self, d):
        '''Multiply all values of dict which are numbers by pointsPerCm
        '''
        from numbers import Number
        from collections import Mapping
        for k, v in d.items():
            if isinstance(v, Mapping):
                self.dictValuesFromCentimeterToPoint(v)
            elif isinstance(v, Number):
                d[k] = v * self.pointsPerCm
        return d

    def drawcentral(self,):
        width = self.sizes['central']['width']
        height = self.sizes['central']['height']
        box = self.slide.Shapes.AddTextbox(
            Orientation=1,  # msoTextOrientationHorizontal
            Left=self.centralCoords['x']-width/2.0,
            Top=self.centralCoords['y']-height/2.0,
            Width=width,
            Height=height,
            )
        box.TextFrame.TextRange.Text = self.centralLabel
        box.TextEffect.Alignment = 2    # centered
        box.Line.Weight = 3
        box.Line.ForeColor.RGB = self.hexToInt('000000')
        self.centralShape = box

    def setradiusKids(self, radius):
        self.radiusKids = float(radius) * self.pointsPerCm

    def setradiusLabels(self, radius):
        self.radiusLabels = float(radius) * self.pointsPerCm

    def drawKids(self,):
        width = self.sizes['kids']['width']
        height = self.sizes['kids']['height']

        for counter, kidKey in enumerate(self.interactionKeys):
            center = self.calcCenter(counter, self.radiusKids)
            box = self.slide.Shapes.AddTextbox(
                Orientation=1,  # msoTextOrientationHorizontal
                Left=center['x']-width/2.0,
                Top=center['y']-height/2.0,
                Width=width,
                Height=height,
                )
            box.TextFrame.TextRange.Text = kidKey
            box.TextEffect.Alignment = 2    # centered
            box.Line.Weight = 3
            box.Line.ForeColor.RGB = self.hexToInt('000000')
            self.kidShape[kidKey] = box

    def xyByPolar(self, radius, angle):
        angle = np.deg2rad(angle)
        pos = {}
        pos['x'] = radius * np.cos(angle)
        pos['y'] = radius * np.sin(angle)
        return pos

    def calcCenter(self, counter, radius):
        nbrKids = len(self.interactionKeys)
        angleIncrement = 360.0 / float(nbrKids)
        angle = counter * angleIncrement
        offset = self.xyByPolar(radius, angle)
        center = {}
        center['x'] = self.centralCoords['x'] + offset['x']
        center['y'] = self.centralCoords['y'] + offset['y']
        return center

    # def drawArrows(self,):
        # for counter, kidKey in enumerate(self.interactionKeys):
            # centerKid = self.calcCenter(counter, self.radiusKids)
            # kidValue = self.interactions[kidKey]
            # arrow = self.slide.Shapes.AddLine(
                # BeginX=self.centralCoords['x'],
                # BeginY=self.centralCoords['y'],
                # EndX=centerKid['x'],
                # EndY=centerKid['y'],
                # ).line
            # arrow.ForeColor.RGB = self.hexToInt('000000')

    def hexToInt(self, hexString):
        # Choose Hex-color: https://www.rapidtables.com/web/color/RGB_Color.html
        h = hexString.lstrip('#')
        rgb = tuple(int(h[i:i+2], 16) for i in (0, 2 ,4))
        colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
        return colorInt

    def drawLabels(self,):
        width = self.sizes['labels']['width']
        height = self.sizes['labels']['height']
        for counter, kidKey in enumerate(self.interactionKeys):
            center = self.calcCenter(counter, self.radiusLabels)
            box = self.slide.Shapes.AddTextbox(
                Orientation=1,  # msoTextOrientationHorizontal
                Left=center['x']-width/2.0,
                Top=center['y']-height/2.0,
                Width=width,
                Height=height,
                )
            box.TextFrame.TextRange.Text = self.formatLabels(kidKey)
            box.TextEffect.Alignment = 2    # centered
            box.TextEffect.FontSize = 10    # centered
            box.Line.Weight = 3
            box.Line.ForeColor.RGB = self.hexToInt('000000')
            box.Fill.BackColor.RGB = self.hexToInt('FFFFFF')
            self.labelShape[kidKey] = box

    def formatLabels(self, kidKey):
        labels = self.interactions[kidKey]
        out = ''
        formats = {
            'labelWith':'Cooperation: ',
            'labelFrom':'Input: ',
            'labelTo':'Output: ',
            }
        for labelKey, labelName in formats.items():
            try:
                out = out + labelName + getattr(labels,labelKey) + '\n'
            except:
                pass
        # Remove trailing newline
        out = out.strip()
        return out

    def connect(self,):
        for counter, kidKey in enumerate(self.interactionKeys):
            pairs = [
                    # [ self.labelShape[kidKey],  self.centralShape ],
                    # [ self.labelShape[kidKey],  self.kidShape[kidKey] ],
                    [ self.centralShape,  self.kidShape[kidKey] ],
                    ]
            for start, end in pairs:
                connector = self.slide.Shapes.AddConnector(
                    Type=1, # msoConnectorStraight
                    BeginX=0,
                    BeginY=0,
                    EndX=0,
                    EndY=0
                    )
                connector.ConnectorFormat.BeginConnect(
                    ConnectedShape=start,
                    ConnectionSite=1,
                    )
                connector.ConnectorFormat.EndConnect(
                    ConnectedShape=end,
                    ConnectionSite=1,
                    )
                connector.ConnectorFormat.Parent.RerouteConnections()
                connector.Line.ForeColor.RGB = self.hexToInt('000000')
                connector.Line.BeginArrowheadStyle = 3
                connector.Line.BeginArrowheadLength = 2
                connector.Line.EndArrowheadStyle = 3
                connector.Line.EndArrowheadLength = 2


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
layoutSlide = p.slides[0].CustomLayout

for keyInteract, valueInteract in interactions.items():
    print(keyInteract)
    g = ConnectionGraph(
            slide=p.slides.AddSlide(1, layoutSlide),
            # Layouts: https://docs.microsoft.com/de-de/office/vba/api/PowerPoint.PpSlideLayout
            centralLabel=keyInteract,
            centralCoords={'x':slideWidth/2.0, 'y':slideHeight/2.0},
            interactions=valueInteract,
            sizesDict_centralKids_WidthHeight={
                'central':{
                    'width':2,  # cm
                    'height':1, # cm
                    },
                'kids':{
                    'width':2,  # cm
                    'height':1, # cm
                    },
                'labels':{
                    'width':4,  # cm
                    'height':1, # cm
                    },
                },
            )
    g.drawcentral()
    g.setradiusKids(radius = 10.0)  # cm
    g.setradiusLabels(radius = 5.0)  # cm
    g.drawKids()
    g.connect()
    g.drawLabels()


# p.Save()
p.SaveAs(filePath+'test')
p.Close()
app.Quit()