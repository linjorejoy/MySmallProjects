arr = "linjo"# [1,2,3,4,5]
arr2 = "linjo" #['a','s','d','f','g']
arr3 = "linjo"
arr4 = "linjo"
arr5 = "linjo"
# [print(a,b,c) for a in arr for b in arr2 for c in arr]
[print(a,b,c,d,e) for a in arr for b in arr2
                     for c in arr3 for d in arr4
                         for e in arr5
                            if not (a  in (b,c,d,e) or
                                b  in (a,c,d,e) or
                                c  in (a,b,d,e) or
                                d  in (a,b,c,e) or
                                e  in (a,b,c,d))]

