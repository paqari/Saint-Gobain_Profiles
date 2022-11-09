import comtypes.client
import math


# create API helper object

helper = comtypes.client.CreateObject('SAP2000v1.Helper').QueryInterface(comtypes.gen.SAP2000v1.cHelper)


mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")

if not mySapObject:
    mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
    print("ok")

# start SAP2000 application
mySapObject.ApplicationStart()

# create SapModel object
SapModel = mySapObject.SapModel

# initialize model
SapModel.InitializeNewModel()

# create new blank model
ret = SapModel.File.NewBlank()

#switch to k-ft units
kgf_m_C = 8
ret = SapModel.SetPresentUnits(kgf_m_C)

#add ASTM A36 material property in United states Region
ret = SapModel.PropMaterial.AddMaterial("A36", 1, "United States", "ASTM A36", "Grade 36")


#add ASTM A653SQGr60
ret = SapModel.PropMaterial.AddQuick("A653SQGr60", 5, 2)


#import new frame section property
# route = r'C:\\Program Files\\Computers and Structures\\SAP2000 23\\Property Libraries\\Sections\\AISC15.xml'
# ret = SapModel.PropFrame.ImportProp('W4X13', 'A36', route, 'W4X13')
# print(ret)

#add coldformed channel
ret = SapModel.PropFrame.SetColdC("C38x38x0.45mm", "A653SQGr50 ", 0.038, 0.038, 0.00045, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdC("C38x38x0.85mm", "A653SQGr50 ", 0.038, 0.038, 0.00085, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdC("C64x38x0.45mm", "A653SQGr50 ", 0.064, 0.038, 0.00045, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdC("C64x38x0.85mm", "A653SQGr50 ", 0.064, 0.038, 0.00085, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdC("C89x38x0.45mm", "A653SQGr50 ", 0.089, 0.038, 0.00045, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdC("C89x38x0.85mm", "A653SQGr50 ", 0.089, 0.038, 0.00085, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdC("C120x38x0.45mm", "A653SQGr50 ", 0.12, 0.038, 0.00045, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdC("C120x38x0.85mm", "A653SQGr50 ", 0.12, 0.038, 0.00085, 0.002, 0.007)

#add coldformed I frame
ret = SapModel.PropFrame.SetColdI("2-C38x38x0.45mm", "A653SQGr50 ", 0.038, 0.064, 0.064, 0.00045, 0.002)
ret = SapModel.PropFrame.SetColdI("2-C38x38x0.85mm", "A653SQGr50 ", 0.038, 0.064, 0.064, 0.00085, 0.002)
ret = SapModel.PropFrame.SetColdI("2-C64x38x0.45mm", "A653SQGr50 ", 0.064, 0.064, 0.064, 0.00045, 0.002)
ret = SapModel.PropFrame.SetColdI("2-C64x38x0.85mm", "A653SQGr50 ", 0.064, 0.064, 0.064, 0.00085, 0.002)
ret = SapModel.PropFrame.SetColdI("2-C89x38x0.45mm", "A653SQGr50 ", 0.089, 0.064, 0.064, 0.00045, 0.002)
ret = SapModel.PropFrame.SetColdI("2-C89x38x0.85mm", "A653SQGr50 ", 0.089, 0.064, 0.064, 0.00085, 0.002)
ret = SapModel.PropFrame.SetColdI("2-C120x38x0.45mm", "A653SQGr50 ", 0.12, 0.064, 0.064, 0.00045, 0.002)
ret = SapModel.PropFrame.SetColdI("2-C120x38x0.85mm", "A653SQGr50 ", 0.12, 0.064, 0.064, 0.00085, 0.002)

#add bracing
ret = SapModel.PropFrame.SetColdC("Bracing-01", "A653SQGr50 ", 0.038, 0.038, 0.00045, 0.002, 0.007)
ret = SapModel.PropFrame.SetColdI("Bracing-02", "A653SQGr50 ", 0.038, 0.064,0.064, 0.00045, 0.002)

#group profile

sec = ["C38x38x0.45mm", "C38x38x0.85mm", "C64x38x0.45mm", "C64x38x0.85mm", "C89x38x0.45mm", "C89x38x0.85mm", "C120x38x0.45mm", "C120x38x0.85mm",
       "2-C38x38x0.45mm", "2-C38x38x0.85mm", "2-C64x38x0.45mm", "2-C64x38x0.85mm", "2-C89x38x0.45mm", "2-C89x38x0.85mm", "2-C120x38x0.45mm", "2-C120x38x0.85mm"]

#draw column
xi=0
yi=0
zi=0

height = [1,2,3,4,5,6]
ret= ' '

for i in sec:
    for j in height:
        [ret, col] = SapModel.FrameObj.AddByCoord(xi, yi, zi, xi, yi, j, ret, i, '1','global')
        xi+=1
    yi+=3
    xi=0









