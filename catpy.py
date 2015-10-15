# notes on the book Python Programming on Win32 by Marc Hammond	

python is written in C
windows uses DLL (dynamic link library) , which allow C code to be stored in one file
and loaded dynamically. Most of pythons functionailty lives in a DLL file

COM(component Object Model) and CORBA(Common Object Request Vroker Architecture)
 allow middle ground between different programming languages to talk to each other on 
 different platforms

VBA
() indicates a function call
[] indicates an array index




#==============================================================================
# CATIA
#==============================================================================
   
#=========================CATIA BOOK===========================================


'''
Book: CATIA V5 MACRO Programming with Visual Basic Script
ISBN: 0071800026
Author: Dieter Ziethen

Interpertor: Neal Gordon
Date: 2014-09-23
Manifesto: In this script, I will attempt to use a VB book to develop commands to be used 
in python

'''

#%% create catia anchor
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
Catia.Visible = True

Catia.Caption
Catia.FullName

Catia.GetWorkBenchID()
Catia.StartWorkBench("PrtCfg")
Catia.StatusBar

#%% pg 23
# only one catia application can be running at a time
# all catia windows can be found using the Windows object
for k in range(Catia.Windows.Count):
    print(Catia.Windows.Item(k+1).Name)

print(Catia.ActiveWindow.Name)

# a list of all documents or files can be generated with this command
for k in range(Catia.Documents.Count):
    print(Catia.Documents.Item(k+1).Name)

print(Catia.ActiveDocument.Name)

# interacting with the OS
Catia.FileSystem.Name
Catia.SystemService.Name
Catia.Printers.Name
Catia.Name

# for each CATIA file, there is a specialized class whos parent class is Document
#   CATPart -> PartDocument , CATProduct -> ProductDocument. The ActiveDocument
#   command automatically detects the type of class

''' pg 25 - Geometry containers in CATPARTS
OriginElements - Origin Planes
AxisSystems - catia name for axis systems
Bodies - partbodies 
HybridBodies - geometric set name
ShapeFactory - solid body and feature creation
HybridShapeFactory - wireframe geometry creation

'''
Part1 = Catia.ActiveDocument
Catia.ActiveDocument.Name

#==============================================================================
# Load, save and close catia documents
#==============================================================================

#%% create catia anchor
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
Catia.Visible = True

#launch CATIA from Python, which executes a batch file, then closes it
import subprocess
filepath=r'C:\Users\ngordon\gdrive\Python\pyscripts\CATIA\Catia_R21_Launch.bat'
p = subprocess.Popen(filepath, creationflags=subprocess.CREATE_NEW_CONSOLE)
# Create an instance of CATIA 
Catia = win32com.client.Dispatch("CATIA.Application")
# close catia
Catia.quit()

# Another Method to open catia, but hold console
import os
os.system(r'C:\Users\ngordon\gdrive\Python\pyscripts\CATIA\Catia_R21_Launch.bat')

# create a new CATIA part
Catia.Documents.Add("Part")
Catia.Documents.Add("Product")
Catia.Documents.Add("Drawing")

#creating a new part from an existing document
# Note - the prefix indicates a raw string and allows the 'newline' \n to be ignored
file1 = r'C:\Users\ngordon\gdrive\Python\pyscripts\CATIA\Part1.CATPart'
Catia.Documents.NewFrom(file1)
# Opening an eisting CATIA document
Catia.Documents.Open(file1)
#saving a CATIA document
Catia.ActiveDocument.Save()
file2 = r'C:\Users\ngordon\gdrive\Python\pyscripts\CATIA\Part2.CATPart'
Catia.ActiveDocument.SaveAs(file2)
#loading a CATIA document, faster than opening but does not show it
Catia.Documents.Read(file1)

# file selection
Catia.FileSelectionBox("File Open","*.CATPart", CATFileSelectionModeOpen)
Catia.FileSelectionBox('File Save','*.CATPart',CATFileSelectionModeSave)

#==============================================================================
# User selection
#==============================================================================

# prints any selection in CATIA
for k in range(Catia.ActiveDocument.Selection.Count):
    print(Catia.ActiveDocument.Selection.Item(k+1).Value.Name)

# selection during runtime
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
Sel1 = Catia.ActiveDocument.Selection
Sel1.Clear
what = ["Pad","Sketch"]
out = Sel1.SelectElement2(what,"make a selection of a pad or a sketch",False)
if out == 'Normal':
    print(Sel1.Item(1).Value.Name)
else:
    print('Selection failed')

#==============================================================================
# CATIA searches, pg 42
#==============================================================================
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
MyList = Catia.ActiveDocument.Selection
MyList.Clear
searchResults = MyList.Search('.Point.Name=Point.1*')


#==============================================================================
# Recognizing Features , pg 43
#==============================================================================
Part1 = Catia.ActiveDocument.Part
Body1 = Part1.MainBody
Geos = Body1.Sketches.Item("Sketch.1").GeometricElements # also use Item(1)
for k in range(Geos.Count):
    eltypnum = Geos.Item(k+1).GeometricType
    if eltypnum == 1:
        eltype = '---axis---'
    elif eltypnum == 2:
        eltype = '---point---'
    elif eltypnum == 3:
        eltype = '---line---'
    elif eltypnum == 5:
        eltype = '---circle---'        
    print(Geos.Item(k+1).Name + ' is a ' + eltype)


#==============================================================================
# Chaning Visibility of parts , Hide/Show, pg 46
#==============================================================================
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
Part1 = Catia.ActiveDocument.Part
Body1 = Part1.MainBody
MyList = Catia.ActiveDocument.Selection
MyList.Clear
# select catia part
MyList.Add(Body1)
MyList.VisProperties.SetShow(True)
MyList.VisProperties.SetShow(False)

#==============================================================================
# Create a new file or declaring an existing file
#==============================================================================
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
# New file 
file3 = r'C:\Users\ngordon\gdrive\Python\pyscripts\CATIA\notes\text_test.txt'
File1 = Catia.FileSystem.CreateFile(file3,False)
# Existing File 
File2 = Catia.FileSystem.GetFile(file3)
File1 == File2

#==============================================================================
# Sends an email from CATIA
#==============================================================================
Email = Catia.CreateMail
Email.SetContent("email sent from catia")
Email.Send



#==============================================================================
#   Inital attempt at figuring out CATIA-COM-Python without the book
#==============================================================================
"""
list of commands used with the COM interface to CATIA
"""

#==============================================================================
#%% First step in every python COM program, launches catia
#==============================================================================
import win32com.client
catia = win32com.client.Dispatch("CATIA.Application")
catia.Visible = True

#==============================================================================
# create a new part, product or drawing
#==============================================================================
part1 = catia.Documents.Add("Part")
prod1 = catia.Documents.Add("Product")
drawing1 = catia.Documents.Add("Drawing")
#drawing1.Close
#catia.Documents.Item("Drawing1").Close



#==============================================================================
# open an existing catia document
#==============================================================================
filepath = "C:\\Users\\ngordon\\gdrive\\Python\\pyscripts\\CATIA\\Part1.CATPart"
catia.Documents.Open(filepath)


#==============================================================================
#%% lanuch catia and open a 'open file' dialog box
#==============================================================================
catia.StartCommand("Open")


#==============================================================================
#%% Part Manipulation
#==============================================================================
#creates a new part in a new window called part1
part1 = catia.Documents.Add("Part").Part
#also
ad = catia.ActiveDocument
part1 = ad.Part

#==============================================================================
#%% Access the Product
#==============================================================================
prod1 = catia.Documents.Add("Product").Product
ad = catia.ActiveDocument
prod1 = ad.Product
prodlist = prod1.Products ; print(prodlist)
#%% lists all parts in a currently open CATProduct
for i in range(prod1.Products.Count):
    print('Part Number:' + prod1.Products.Item(i+1).PartNumber)
    

#==============================================================================
#%% Find the volume of the current body
#==============================================================================
stringer = catia.ActiveDocument.Part
vol = stringer.Analyze.Volume

#==============================================================================
#%% make a cube 
#==============================================================================

import  random

catia.Visible = True

part1 = catia.Documents.Add("Part").Part
ad = catia.ActiveDocument
part1 = ad.Part
bod = part1.MainBody
bod.Name="cube"
cubeWidth = 10

skts = bod.Sketches
xyPlane = part1.CreateReferenceFromGeometry(part1.OriginElements.PlaneXY)
shapeFact = part1.Shapefactory

ms = skts.Add(xyPlane)
ms.Name="Cube Outline"

fact = ms.OpenEdition()
fact.CreateLine(-cubeWidth, -cubeWidth,  cubeWidth, -cubeWidth)
fact.CreateLine(cubeWidth, -cubeWidth,  cubeWidth, cubeWidth)
fact.CreateLine(cubeWidth, cubeWidth,  -cubeWidth, cubeWidth)
fact.CreateLine(-cubeWidth, cubeWidth,  -cubeWidth, -cubeWidth)
ms.CloseEdition()
mpad = shapeFact.AddNewPad(ms, cubeWidth)
mpad.Name = "Python Pad"
mpad.SecondLimit.Dimension.Value = cubeWidth

sel = ad.Selection
sel.Add(mpad)

vp = sel.VisProperties
vp.SetRealColor(random.randint(0,255),random.randint(0,255),random.randint(0,255), 0)
part1.Update()

#==============================================================================
#%% get current product, and return tree information
#==============================================================================
import win32com.client
catia = win32com.client.Dispatch("CATIA.Application")

ad = catia.ActiveDocument
prod1 = ad.Product

# lists all parts in a currently open CATProduct
for i in range(prod1.Products.Count):
    print('Part Number:' + prod1.Products.Item(i+1).PartNumber)

#==============================================================================
#%% Run makepy.py to get early-bound, static ojects, not working!!
#==============================================================================
'''
CATIA V5 CATArrangementInterfaces Object Library
 {A903D4EA-3932-11D3-8BB3-006094EB5532}, lcid=0, major=0, minor=0
 >>> # Use these commands in Python code to auto generate .py support
 >>> from win32com.client import gencache
 >>> gencache.EnsureModule('{A903D4EA-3932-11D3-8BB3-006094EB5532}', 0, 0, 0)
'''
import win32com.client
win32com.client.gencache.EnsureModule('{A903D4EA-3932-11D3-8BB3-006094EB5532}', 0, 0, 0)
catia = win32com.client.Dispatch("CATIA.Application")

#==============================================================================
#%% test
#==============================================================================
import win32com.client

# force early binding
#win32com.client.gencache.EnsureModule('{A903D4EA-3932-11D3-8BB3-006094EB5532}', 0, 0, 0)
Catia = win32com.client.Dispatch("CATIA.Application")
filepath = 'C:\\Users\\ngordon\\gdrive\\Python\\pyscripts\\CATIA\\stringer-assembly\\V5313904620200_D00_CNC_1.CATProduct'
#filepath = "C:\\Users\\ngordon\\gdrive\\Python\\pyscripts\\CATIA\\Part1.CATPart"
Catia.Documents.Open(filepath)

Catia.ActiveDocument.Name

Catia.ActiveDocument.Product.GetActiveShapeName

Catia.ActiveDocument.Product.Products.AddNewProduct("ass1")


Catia.ActiveDocument.Product.Products.Count

Catia.ActiveDocument.Product.Products.Item(1).PartNumber

Catia.ActiveDocument.Product

Catia.ActiveDocument.Product.Products.Item(1)

doc1 = Catia.ActiveDocument
doc1

prods = Catia.ActiveDocument.Product.Products
for k in prods:
    print(k.Name)

part1 = catia.ActiveDocument.Product.Products.Item(1).MainBody

Catia.ActiveDocument.Close

AddComponentsFromFiles(cn_list, "All")

#==============================================================================
#%% BoltVBS
#==============================================================================
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
MyDocument = Catia.Documents.Add("Part")
print(MyDocument.Name)
PartFactory = MyDocument.Part.ShapeFactory  # Retrieve the Part Factory.
MyBody1 = MyDocument.Part.Bodies.Item("PartBody")
Catia.ActiveDocument.Part.InWorkObject = MyBody1 # Activate "PartDesign#

# Creating the Shaft
ReferencePlane1 = MyDocument.Part.CreateReferenceFromGeometry(MyDocument.Part.OriginElements.PlaneYZ)
  
#Create the sketch1 on ReferencePlane1
Sketch1 = MyBody1.Sketches.Add(ReferencePlane1)
MyFactory1 = Sketch1.OpenEdition() # Define the sketch

h1 = 80 # height of the bolt
h2 = 300 # total height
r1 = 120 # external radius
r2 = 60 # Internal radius
s1 = 20 # Size of the chamfer
  
l101 = MyFactory1.CreateLine(0, 0, r1 - 20, 0)
l102 = MyFactory1.CreateLine(r1 - 20, 0, r1, -20)
l103 = MyFactory1.CreateLine(r1, -20, r1, -h1 + 20)
l104 = MyFactory1.CreateLine(r1, -h1 + 20, r1 - 20, -h1)
l105 = MyFactory1.CreateLine(r1 - 20, -h1, r2, -h1)
l106 = MyFactory1.CreateLine(r2, -h1, r2, -h2 + s1)
l107 = MyFactory1.CreateLine(r2, -h2 + s1, r2 - s1, -h2)
l108 = MyFactory1.CreateLine(r2 - s1, -h2, 0, -h2)
l109 = MyFactory1.CreateLine(0, -h2, 0, 0)
Sketch1.CenterLine = l109
  
Sketch1.CloseEdition
AxisPad1 = PartFactory.AddNewShaft(Sketch1)
  
#' Creating the Pocket
ReferencePlane2 = MyDocument.Part.CreateReferenceFromGeometry(MyDocument.Part.OriginElements.PlaneXY)
    
# Create the sketch2 on ReferencePlane2
Sketch2 = MyBody1.Sketches.Add(ReferencePlane2)
MyFactory2 = Sketch2.OpenEdition()
D = 1 / 0.866
  
l201 = MyFactory2.CreateLine(D * 100, 0, D * 50, D * 86.6)
l202 = MyFactory2.CreateLine(D * 50, D * 86.6, D * -50, D * 86.6)
l203 = MyFactory2.CreateLine(D * -50, D * 86.6, D * -100, 0)
l204 = MyFactory2.CreateLine(D * -100, 0, D * -50, D * -86.6)
l205 = MyFactory2.CreateLine(D * -50, D * -86.6, D * 50, D * -86.6)
l206 = MyFactory2.CreateLine(D * 50, D * -86.6, D * 100, 0)

#  ' Create a big circle around the form to get a Hole
c2 = MyFactory2.CreateClosedCircle(0, 0, 300)
  
Sketch2.CloseEdition
AxisHole2 = PartFactory.AddNewPocket(Sketch2, h1)
  
MyDocument.Part.Update


#==============================================================================
#%% CATIA V5: Macro Programming Dieter Ziethen pg 39 open 
#==============================================================================
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
Catia.Visible = True
Catia.Documents.NewFrom("C:\\Users\\ngordon\\gdrive\\Python\\pyscripts\\CATIA\Part1.CATPart")

#open from 
#Catia.StartCommand("Open")


#==============================================================================
#%% CATIA V5: Macro Programming Dieter Ziethen pg 39, user 
#==============================================================================
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")

userselect = Catia.ActiveDocument.Selection
for k in range(userselect.Count):
    print(userselect.Item(k+1).Value.Name)

#==============================================================================
#%% CATIA V5: Macro Programming Dieter Ziethen pg 39, searching
#=======================================================recognizing geometric elements=======================
'''
CATIA V5: Macro Programming with Visual Basic Script Dieter Ziethen
pg 45 , searching
syntax is as follows
[Environment].[Type].[Attribute]=[Value];[Search]
Catia Workbench = enviroment# param_tree
for k in range(part1.Parameters.Count):
    print(part1.Parameters.Item(k+1).Name)
Type = element type

'''
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
mysel = Catia.ActiveDocument.Selection # recognizing geometric elements
mysel.Search(".Point.Name=Point.1*;All")


#==============================================================================
#%% CATIA V5: Macro Programming Dieter Ziethen pg 39, user selection,recognizing geometric elements
#==============================================================================

import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")
part = Catia.ActiveDocument.Part
body = part.MainBody
#instead of explicitly calling "Sketch.1", you can use a numeric value
geoset = body.Sketches.Item("Sketch.1").GeometricElements

for k in range(geoset.Count):
    print(geoset.Item(k+1).Name)

#==============================================================================
#%% My code for digging in geosets
#==============================================================================

'''
# one method by defining a bunch of objects
documents1 = Catia.Documents
partDocument1 = documents1.Item("V5313904620200_D00_CG_1.CATPart")
part1 = partDocument1.Part
hybridBodies1 = part1.HybridBodies
hybridBody1 = hybridBodies1.Add()
hybridBodies2 = hybridBody1.HybridBodies
hybridBody2 = hybridBodies2.Add()

# another way keeps it clearer
Catia.Documents.Item("V5313904620200_D00_CG_1.CATPart").Part.HybridBodies.Add().HybridBodies.Add()
'''
''' GEOSET CREATION EXAMPLE
#k sub geosets
part1 = Catia.Documents.Item("V5313904620200_D00_CG_1.CATPart").Part
for k in range(5):
    part1 = part1.HybridBodies.Add()
    part1.Name = 'geoset_'+str(k)


g = 5  # number of main geosets
sg = 4 # number of subgeosets
#k Main geosets
for j in range(g):
    part1 = Catia.Documents.Item("V5313904620200_D00_CG_1.CATPart").Part
    part1 = part1.HybridBodies.Add()
    part1.Name = 'Main_geoset_'+str(j)
    # subgeoset
    for k in range(sg):
        part1 = part1.HybridBodies.Add()
        part1.Name = 'subgeoset_'+str(j)+'_'+str(k)
'''

# manual geoset dig 0
part1.HybridBodies.Item(1).HybridBodies.Item(1).Name

# manual geoset dig 1
for k in range(part1.HybridBodies.Count):
    print(part1.HybridBodies.Item(k+1).Name)
    if part1.HybridBodies.Item(k+1).Name == "WJ INPUTS":
        print('found it!!!!')
    
# manual geoset dig 2
part1 = Catia.Documents.Item("V5313904620200_D00_CG_1.CATPart").Part
for k in range(part1.HybridBodies.Count):
    print(part1.HybridBodies.Item(k+1).Name)
    for j in range(part1.HybridBodies.Item(k+1).HybridBodies.Count ):
        print('   '+part1.HybridBodies.Item(k+1).HybridBodies.Item(j+1).Name)


#part1.HybridBodies.Item(3).HybridBodies.Count 


hybridShapeCurveExplicit1 = part1.Parameters.Item("Hole-1")
part1.HybridBodies.Item("Hole_Geometry").HybridBodies.Item("Starboard").HybridBodies.Item("hole_1").Parameters



# param_tree
for k in range(part1.Parameters.Count):
    print(part1.Parameters.Item(k+1).Name)
    

#==============================================================================
#  printout of all geometric set geometry
#==============================================================================
def geoset_tree(part1,indent = ' '):
    for k in range(part1.HybridBodies.Count):
        print(indent+part1.HybridBodies.Item(k+1).Name)
        if part1.HybridBodies.Count > 1:
            geoset_tree(part1.HybridBodies.Item(k+1),indent+'   ')

geoset_tree(part1)


#==============================================================================
# CATProcess Manipulation
#==============================================================================
import win32com.client
Catia = win32com.client.Dispatch("CATIA.Application")

# load file
file1 = r'X:\Work Directory\ngordon\V5332067420000\V5332067420000_A00.CATProcess'
Catia.Documents.NewFrom(file1)
Catia.Visible = True
#Catia.Documents.Read(file1)

# Show some information on the file
for k in range(Catia.Windows.Count):
    print(Catia.Windows.Item(k+1).Name)
print(Catia.ActiveWindow.Name)
# a list of all documents or files can be generated with this command
for k in range(Catia.Documents.Count):
    print(Catia.Documents.Item(k+1).Name)
# File name
print(Catia.ActiveDocument.Name)

# prints any selection in CATIA
for k in range(Catia.ActiveDocument.Selection.Count):
    print(Catia.ActiveDocument.Selection.Item(k+1).Value.Name)
    
# Prints ALL the process parameters in the file
param_process = Catia.ActiveDocument.GetItem("Process").Parameters
for k in range(param_process.Count):
    print(param_process.Item(k+1).Name)

s = r'ROUGH OUTER FLANGE 1\MfgParameter.5\Max lead angle'
param_process.Item(s).Value = 25

s = r'Process\ROUGH RH END\MfgParameter.5\Percentage overlap'
param_process.Item(s).Value = 12

s = r'ROUGH RH END\MfgParameter.2\Distance after corner'
param_process.Item(s).Value = 2.0

s = 'OUTER FLANGE RELIEF CUT 5\MfgParameter.2\MFG_FEED_SELECT_TRANSITION'
param_process.Item(s).Value = True

s = r'Process\OUTER FLANGE RELIEF CUT 5\MfgParameter.5\Max discretization step'
param_process.Item(s).Value = 10006

s = r'Process\OUTER FLANGE RELIEF CUT 5\MfgParameter.5\Type of contouring'
param_process.Item(s).Value = 'Angular'

param_process.Parent.Name

pprActivity = Catia.ActiveDocument.GetItem("Process")
pprActivity.PPRItems
pprActivity.PPRItems.Name
pprActivity.Description
pprActivity.Name
pprActivity.Type
pprActivity.Items.Count
pprActivity.Parent.Name
pprActivity.Parameters.Name



# P.P.R -> ProcessList -> Process -> Attributes -> Parameters
Catia.ActiveDocument.GetItem("Process").Parameters.Item("ROUGH MACHINING TOLERANCE").Value = .01

Catia.Documents.Item("V5332067420000_A00_CG.CATPart")

#Commands
Catia.StartCommand('Fit All In')

Catia.ActiveDocument.Product
Analyze(Catia.ActiveDocument.Product)

