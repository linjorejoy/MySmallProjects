# Importing Image from PIL package  
from PIL import Image
import pandas as pd
import numpy as np
from numpy import asarray

# creating a image object 
colourImg = Image.open(r"G:/my works new/Python/Image to excel/Screenshot (241).png")
colourPixels = colourImg.convert("RGB")


colourArray = np.array(colourPixels.getdata()).reshape(colourImg.size + (3,))
indicesArray = np.moveaxis(np.indices(colourImg.size), 0, 2)
allArray = np.dstack((indicesArray, colourArray)).reshape((-1, 5))

df = pd.DataFrame(allArray, columns=["y", "x", "red","green","blue"])

px = colourPixels.load()
print(colourPixels.size)
print(colourPixels.mode)

#numpy_data = np.array([])
#data= asarray(colourPixels)
#print(data.shape)

#print(data)

#here first [] is for row then next [] is for column next [] is for rgba
#print(data[0][1][0])


print (df['red'][2])
print (df)


#for i in range (0,im.size[0]):
    #for j in range (0,im.size[1]):
        #for k in range(0,3):
            #print (px[i, j][0],"  ",px[i, j][1],"  ",px[i, j][2])
            #numpy_data = np.append(numpy_data,px[i, j][k])
        


