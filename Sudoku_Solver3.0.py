grid= [[5,3,0,0,7,0,0,0,0],
       [6,0,0,1,9,5,0,0,0],
       [0,9,8,0,0,0,0,6,0],
       [8,0,0,0,6,0,0,0,3],
       [4,0,0,8,0,3,0,0,1],
       [7,0,0,0,2,0,0,0,6],
       [0,6,0,0,0,0,2,8,0],
       [0,0,0,4,1,9,0,0,5],
       [0,0,0,0,8,0,0,0,0]]
import numpy as np
print(np.matrix(grid))

count = 0
count2 = 0

def counter():
    global count
    count+=1
    if (count %10==0):
        print(" count = " , count)
        return count
    else:
        pass
    

def counter2():
    global count2
    count2+=1
    if (count2 %10==0):
        print(" count2 = " , count2)
        return count2
    else:
        pass

def possible(y,x,n):
    global grid
    for i in range(0,9):
        if grid[y][i] == n:
            return False
    for i in range(0,9):
        if grid[i][x] == n:
            return False
    x0=(x//3)*3
    y0=(y//3)*3
    for i in range(0,3):
        for j in range(0,3):
            if grid[y0+i][x0+j] == n:
                return False
    return True

def solve():
    global grid
    for y in range(9):
        for x in range(9):
            if grid[y][x] == 0:
                for n in range(1,10):
                    if possible(y,x,n):
                        counter()
                        grid[y][x]= n
                        solve()
                        counter2()
                        grid[y][x]= 0
                return
            
    print(np.matrix(grid))
    input("More?")
