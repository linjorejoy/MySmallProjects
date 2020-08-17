import numpy as np
import math
import mathplotlib.pyplot as plt
stp= input (print("step division = "))
time = np.arrange(0,2*math.pi,math.pi/stp)

amplitude = np.sin(time)
plt.plot(time,amplitude)

plt.show()
