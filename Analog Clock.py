import turtle
wn = turtle.Screen()
wn.bgcolor("white")
wn.setup(width=600,height=600)
wn.title("Clock")

#drawing

pen = turtle.Turtle()
pen.hideturtle()
pen.speed(0)
pen.pensize(3)
 

def dr_cl(pen):
    pen.up()
    pen.goto(0,210)
    pen.setheading(180)
    pen.color("black")
    pen.pendown()
    pen.circle(210)


    #hour
    pen.penup()
    pen.goto(0,0)
    pen.setheading(90)


    for i in range(12):
        pen.fd(190)
        pen.pendown()
        pen.fd(20)
        pen.penup()
        pen.goto(0,0)
        pen.rt(30)
    

dr_cl(pen)

wn.mainloop()
