import os

files = os.listdir('processed/')

try:
    os.mkdir('sliced')
except:
    pass

for file in files:
    with open('processed/'+file) as f:
        data = f.read()
        dt = data.split()
        a,b,c,d,e,f,g,h,i,j,k,l = None,None,None,None,None,None,None,None,None,None,None,None
        a = ' '.join(dt[:len(dt)//12])
        with open('sliced/'+file[:-4]+' p1'+'.txt', 'w') as f:
            f.write(a)
        b = ' '.join(dt[len(dt)//12:((len(dt)//12)*2)])
        with open('sliced/'+file[:-4]+' p2'+'.txt', 'w') as f:
            f.write(b)
        c = ' '.join(dt[((len(dt)//12)*2):((len(dt)//12)*3)])
        with open('sliced/'+file[:-4]+' p3'+'.txt', 'w') as f:
            f.write(c)
        d = ' '.join(dt[((len(dt)//12)*3):((len(dt)//12)*4)])
        with open('sliced/'+file[:-4]+' p4'+'.txt', 'w') as f:
            f.write(d)
        e = ' '.join(dt[((len(dt)//12)*4):((len(dt)//12)*5)])
        with open('sliced/'+file[:-4]+' p5'+'.txt', 'w') as f:
            f.write(e)
        f = ' '.join(dt[((len(dt)//12)*5):((len(dt)//12)*6)])
        with open('sliced/'+file[:-4]+' p6'+'.txt', 'w') as z:
            z.write(f)
        g = ' '.join(dt[((len(dt)//12)*6):((len(dt)//12)*7)])
        with open('sliced/'+file[:-4]+' p7'+'.txt', 'w') as f:
            f.write(g)
        h = ' '.join(dt[((len(dt)//12)*7):((len(dt)//12)*8)])
        with open('sliced/'+file[:-4]+' p8'+'.txt', 'w') as f:
            f.write(h)
        i = ' '.join(dt[((len(dt)//12)*8):((len(dt)//12)*9)])
        with open('sliced/'+file[:-4]+' p9'+'.txt', 'w') as f:
            f.write(i)
        j = ' '.join(dt[((len(dt)//12)*9):((len(dt)//12)*10)])
        with open('sliced/'+file[:-4]+' p10'+'.txt', 'w') as f:
            f.write(j)
        k = ' '.join(dt[((len(dt)//12)*10):((len(dt)//12)*11)])
        with open('sliced/'+file[:-4]+' p11'+'.txt', 'w') as f:
            f.write(k)
        l = ' '.join(dt[((len(dt)//12)*11):])
        with open('sliced/'+file[:-4]+' p12'+'.txt', 'w') as f:
            f.write(l)
