import pandas


with open("ocldb1622055189.22021.CTD.csv", 'r') as file:
    content = file.readlines()

lst = []
big = []
for line in content:
    if "END OF VARIABLES SECTION," in line:
        lst.append(line)
        big.append(lst)
        lst = []
    else:
        lst.append(line)
x = 0
for item in big:
    check = item[-2].split(",")
    if float(check[1].strip()) >= 500.0:
        x+=1
        with open("new1.csv" ,"a") as outfile:
            for em in item:
                outfile.write(em)
print(x)



    # print(item[-2])
