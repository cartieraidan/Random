dict = {}

with open("dictKeyValues.txt", "r") as file:
    tmp = ""
    for line in file:
        
        tmp = line.split('"')

        dict[tmp[3]] = tmp[1]

print(dict)