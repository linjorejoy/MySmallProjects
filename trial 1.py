# con="y"
# while(con=="y"):
#     num = int(input("enter the number : "))
#     for i in range(1,21):
#         print(num,"x",i,"=",num*i)
#     con = input (print("continue y/n "))
    
class PrimeGenerator:
    def __init__(self, maximum):
        self.maximum = maximum + 1
        self.count = 0

    def __next__(self):
        if self.count < self.maximum:
            for num in range(self.count, self.maximum):
                for divisor in range(2, self.count + 1):
                    if num // divisor != num / divisor:
                        print(f'number is prime, it is {num} and the count is {self.count}, and the divisor is {divisor}')
                        self.count += 1
                        return num
                    else:
                        break

                self.count += 1
        else:
            return

my_gen = PrimeGenerator(7)
print(next(my_gen))
print(next(my_gen))
print(next(my_gen))
print(next(my_gen))
print(next(my_gen))