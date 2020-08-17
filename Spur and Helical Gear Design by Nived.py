# program for design of gears
# here first we have to choose which gear we have to design

print("Author : Nived N")
print("Email ID: nived00015@gmail.com")

print("Note this program is for designing spur and helical gears only")
print("While typing gear you want to design, please enter either 'spur' or 'helical'")

print("")
print("")
print("")
print("")
print("")
import math
gear = input("type of gear you want to design:")
spur_helical_modules = [1,1.25,1.5,2,2.5,3,4,5,6,8,10,12,16,20,25,32,40,50]

if gear == "spur":
    # input all values required for calculation.
    print("Enter the number of teeths in pinion (if not given, assume)")
    Zp = float(input("Enter the number of teeth on pinion:"))
    i = float(input("Enter the gear ratio:"))
    Zg = i*Zp
    print("The number of teeth on gear is",Zg)
    Spd = input("Enter the name of material used for pinion:")
    Sp = float(input("Enter the allowable static stress of pinion:"))
    Sgd = input("Enter the name of material used for gear:")
    Sg = float(input("Enter the allowable static stress of gear:"))
    P = float(input("Enter the power (in kW):"))
    d = float(input("Enter the pinion diameter (if not given give 0 value):"))
    D = float(input("Enter the Gear diameter (if not given give 0 value):"))
    Np = float(input("Enter the Speed of pinion (if not given give 0 value):"))
    Ng = float(input("Enter the speed of Gear (if not given give 0 value):"))
    print(" ")
    
    k = float(input("Enter the width-module ratio:"))
    Cs = float(input("Enter Service Factor (take it from table 12.8, page no:235):"))
    
 
    
    
    
    # using louis form factor calculate weaker part and required calculation prior to module calculation
    
    teeth_type = input("Enter the type of teeth (20 degree stub teeth, 20 degree involute, 14.5 degree involute):")
    if teeth_type == '20 degree involute':
        yp = (0.154)-(0.912/Zp)
        yg = (0.154)-(0.912/Zg)
    if teeth_type == '20 degree stub teeth':
        yp = (0.175)-(0.95/Zp)
        yg = (0.175)-(0.95/Zg)

    if teeth_type == '14.5 degree involute':
        yp = (0.124)-(0.684/Zp)
        yg = (0.124)-(0.684/Zg)
    a = Sp*yp
    b=  Sg*yg
    if a<b:
        print("Pinion is weaker, So design is based on pinion")
        T = (P*(10**6)*60)/(2*3.14*Np)
        v = (3.14*d*Np)/60000
        Yp = 3.14*yp
        
    else:
        print("Gear is weaker, so design is based on gear")
        T = (P*(10**6)*60)/(2*3.14*Ng)
        v = (3.14*D*Ng)/60000
        Yg = 3.14*yg
        

    # calculate velocity factor

    print("The velocity is",v,"m/s")

    print("gear types: ordinary cut gears, carefully cut gears, accurately cut gears, precision gears, non mettalic gears, unknown (Check out velocity factor table of spur gear)")
    gear_type = input("Enter gear type as shown above:")
    
    
    if v <=8 and gear_type == 'ordinary cut gears':
          Cv = 3.05/(3.05+v)
    if v<=13 and gear_type =='carefully cut gears':
          Cv = 4.58/(4.58+v)
    if 6<v<=20 and gear_type =='accurately cut gears':
          Cv = 6.1/(6.1+v)
    if v>20 and gear_type == 'precision gears':
          a1 = math.sqrt(v)
          Cv = 5.55/(5.55+a1)
    if gear_type == 'non mettalic gears':
          Cv = 0.7625/(1.0167+v)

    if v ==0 and gear_type =='unknown':
        Cv =0.5
        
          

    # module calculation
    if a<b:
          cube_module = (2*T)/(Sp*Cv*k*Yp*Zp)
          m = (cube_module)**(1/3)
          print("The calculated value of module is",m)
    else:
          cube_module = (2*T)/(Sg*Cv*k*Yg*Zg)
          m = (cube_module)**(1/3)
          print("The calculated value of module is",m)
    print("Standerdising the module we get......")

    l1 = len(spur_helical_modules)
    for r in range(l1):
          if spur_helical_modules[r]>=m:
              m = spur_helical_modules[r]
              flag = r
              break
              
    Q = D
    q = d
    print("The standerdised value of module is",m)

    # checking safety of module using lewis equation (page no.204, eq.12.5(a))
  
    if a<b:
        if d ==0:
            d = m*Zp
            v = (3.14*d*Np)/60000
        
    if a>b:
        if D == 0:
            D = m*Zg
            v= (3.14*D*Ng)/60000
  
   
    # calculate velocity factor (for safety purposes)

    print("The velocity is",v,"m/s")

    print("gear types: ordinary cut gears, carefully cut gears, accurately cut gears, precision gears, non mettalic gears (Check out velocity factor table of spur gear)")
    gear_type = input("Enter gear type as shown above:")
    
    if v <=8 and gear_type == 'ordinary cut gears':
          Cv = 3.05/(3.05+v)
    if v<=13 and gear_type =='carefully cut gears':
          Cv = 4.58/(4.58+v)
    if 6<v<=20 and gear_type =='accurately cut gears':
          Cv = 6.1/(6.1+v)
    if v>20 and gear_type == 'precision gears':
          a1 = math.sqrt(v)
          Cv = 5.55/(5.55+a1)
    if gear_type == 'non mettalic gears':
          Cv = 0.7625/(1.0167+v)

    # Calculate tangential tooth load
    Ft = (1000*P*Cs)/(v)
    b1 = k*m
    if a<b:
        Sp1 = Ft/(Cv*b1*Yp*m)
    if a>b:
        Sg1 = Ft/(Cv*b1*Yg*m)
        

    if a<b:
        if Sp1<Sp:
            print("")

        else:
            while Sp1>Sp:
                flag = flag+1
                m = spur_helical_modules[flag+1]
                Sp1 = Ft/(Cv*b1*Yp*m)

    print("the gear is safe with a module value of",m)

    # calculate dimension
    
    
    if Q ==0:
        D = m*Zg
    if q ==0:
        d = m*Zp
    if Q !=0:
        assume = input("Do no of teeth in pinion were assumed?:")
        if assume == "No":
            print("")
        else:
            Zp = d/m
            Zg = D/m
    b1 = k*m

    # print dimensions
    print("The Final-Fixed dimensions of pinion-gear")
    print("")
    print("The number of teeth on pinion: ",Zp)
    print("")
    print("The number of teeth on gear: ",Zg)
    print("")
    print("The Pitch circle Dia of pinion: ",d,"mm")
    print("")
    print("The Pitch circle Dia of Gear: ",D,"mm")
    print("")
    print("The module:",m)
    print("")
    print("The face width",b1,"mm")
    print("")
    print("The velocityis:",v,"m/s")
    


    # force analysis

    # calculate dynamic load
    K3 = 20.67
    C = float(input("Enter the dynamic factor (used for dynamic load calculations):"))
    numerator = (K3*v)*((C*b1)+Ft)
    L = math.sqrt((C*b1)+Ft)
    denominator = (K3*v)+L
    Fi = numerator/denominator
    Fd = Ft+Fi
    print("The Dynamic load is:",Fd,"N")

    # Calculate endurance strength


    # calculate wear load

    j = (2*Zg)/(Zp+Zg)
    K = float(input("Enter load stress factor:"))

    Fw = d*b1*j*K
    print("The Wear load is:",Fw,"N")
    if Fw>=Fd:
        print("The gear is safe against Wear load")
    else:
        w = Fw/(d*b1*j)
        print("Materials sets having load stress factor greater than",w,"should be taken")


if gear == "helical":
    # input all values required for calculation.
    print("Enter the number of teeths in pinion (if not given, assume)")
    Zp = float(input("Enter the number of teeth on pinion:"))
    i = float(input("Enter the gear ratio:"))
    Zg = i*Zp
    print("The number of teeth on gear is",Zg)
    Spd = input("Enter the name of material used for pinion:")
    Sp = float(input("Enter the allowable static stress of pinion:"))
    Sgd = input("Enter the name of material used for gear:")
    Sg = float(input("Enter the allowable static stress of gear:"))
    P = float(input("Enter the power (in kW):"))
    d = float(input("Enter the pinion diameter (if not given give 0 value):"))
    D = float(input("Enter the Gear diameter (if not given give 0 value):"))
    Np = float(input("Enter the Speed of pinion (if not given give 0 value):"))
    Ng = float(input("Enter the speed of Gear (if not given give 0 value):"))
    bt = float(input("Enter the helix angle in degrees:"))
    k = float(input("Enter the width-module ratio:"))
    Cs = float(input("Enter Service Factor (take it from table 12.8, page no:235):"))
    Cw = float(input("Enter the wear and lubrication factor:"))
    print("")
    print(" ")
    

     # using louis form factor calculate weaker part and required calculation prior to module calculation

    rad = (3.14*bt)/180
    an = math.cos(rad)
    a1 = an**3
    Zpe = Zp/a1
    Zge = Zg/a1
    
    teeth_type = input("Enter the type of teeth (20 degree stub teeth, 20 degree involute, 14.5 degree involute):")
    if teeth_type == '20 degree involute':
        yp = (0.154)-(0.912/Zpe)
        yg = (0.154)-(0.912/Zge)
    if teeth_type == '20 degree stub teeth':
        yp = (0.175)-(0.95/Zpe)
        yg = (0.175)-(0.95/Zge)

    if teeth_type == '14.5 degree involute':
        yp = (0.124)-(0.684/Zpe)
        yg = (0.124)-(0.684/Zge)
    a = Sp*yp
    b=  Sg*yg

    if a<b:
        print("Pinion is weaker, So design is based on pinion")
        T = (P*(10**6)*60)/(2*3.14*Np)
        v = (3.14*d*Np)/60000
        Yp = 3.14*yp
        
    else:
        print("Gear is weaker, so design is based on gear")
        T = (P*(10**6)*60)/(2*3.14*Ng)
        v = (3.14*D*Ng)/60000
        Yg = 3.14*yg

    # calculate velocity factor
    if v<=5:
        Cv = 4.58/(4.58+v)
    if 5<v<=10:
        Cv = 6.1/(6.1+v)
    if 10<v<=20:
        Cv = 15.25/(15.25+v)
    if v>20:
        ang = math.sqrt(v)
        Cv = 5.55/(5.55+ang)
    if v ==0:
        Cv = 0.5
        
    
    # calculation of module

    
    
    
    
    
    


    if a<b:
          cube_module = (2*T*Cw*an)/(Sp*Cv*k*Yp*Zp)
          mn = (cube_module)**(1/3)
          print("The calculated value of module is",mn)
    else:
          cube_module = (2*T*Cw*an)/(Sg*Cv*k*Yg*Zg)
          mn = (cube_module)**(1/3)
          print("The calculated value of module is",mn)
    print("")
    print("")
    print("Standerdising the module we get......")

    l1 = len(spur_helical_modules)
    for r in range(l1):
          if spur_helical_modules[r]>=mn:
              mn = spur_helical_modules[r]
              flag = r
              break


    Q = D
    q = d
    print("The standerdised value of module is",mn)

    # checking safety of module using lewis equation (page no.204, eq.12.5(a))
  
    if a<b:
        if d ==0:
            m = mn/an       
            d = m*Zp
            v = (3.14*d*Np)/60000
        
    if a>b:
        if D == 0:
            m = mn/an
            D = m*Zg
            v= (3.14*D*Ng)/60000

    


     # calculate velocity factor
    if v<=5:
        Cv = 4.58/(4.58+v)
    if 5<v<=10:
        Cv = 6.1/(6.1+v)
    if 10<v<=20:
        Cv = 15.25/(15.25+v)
    if v>20:
        ang = math.sqrt(v)
        Cv = 5.55/(5.55+ang)

    
    # Calculate tangential tooth load
    Ft = (1000*P*Cs)/(v)
    b1 = k*mn
    
    if a<b:
        Sp1 = Ft/(Cv*b1*Yp*mn)
    if a>b:
        Sg1 = Ft/(Cv*b1*Yg*mn)


    if a<b:
        if Sp1<Sp:
            print("")

        else:
            while Sp1>Sp:
                flag = flag+1
                mn = spur_helical_modules[flag+1]
                Sp1 = Ft/(Cv*b1*Yp*mn)



    print("the gear is safe with a module value of",mn)


     # calculate dimension
    
    
    if Q ==0:
        m = mn/an
        D = m*Zg
    if q ==0:
        m =mn/an
        d = m*Zp
    if Q !=0:
        assume = input("Do no of teeth in pinion were assumed?:")
        if assume == "No":
            print("")
        else:
            
            Zp = d*an/mn
            Zg = D*an/mn

    b1 = k*mn

    # print dimensions
    print("The Final-Fixed dimensions of pinion-gear")
    print("")
    print("The number of teeth on pinion: ",Zp)
    print("")
    print("The number of teeth on gear: ",Zg)
    print("")
    print("The Pitch circle Dia of pinion: ",d,"mm")
    print("")
    print("The Pitch circle Dia of Gear: ",D,"mm")
    print("")
    print("The module:",mn)
    print("")
    print("The face width",b1,"mm")
    print("")
    print("The velocityis:",v,"m/s")


    # force analysis

    # calculate dynamic load
    K3 = 20.67
    C = float(input("Enter the dynamic factor (used for dynamic load calculations):"))
    numerator = (K3*v)*((C*b1*(an**2))+Ft)*(an)
    L = math.sqrt((C*b1*(an**2))+Ft)
    denominator = (K3*v)+L
    Fi = numerator/denominator
    Fd = Ft+Fi
    print("The Dynamic load is:",Fd,"N")

    # Calculate endurance strength


    # calculate wear load

    j = (2*Zg)/(Zp+Zg)
    K = float(input("Enter load stress factor:"))

    Fw = d*b1*j*K/(an**2)
    print("The Wear load is:",Fw,"N")
    if Fw>=Fd:
        print("The gear is safe against Wear load")
    else:
        w = Fw*(an**2)/(d*b1*j)
        print("Materials sets having load stress factor greater than",w,"should be taken")


    
    


    

                

            



    


    


    

    

    

    

    
    

input()
    
    
    



    
    
    
    
    
    
                
                
                
        
    
    
     
    
        
    
    
    
          
          
          
          
          
    
    
    
    

    
    
    

    
    
    
    
