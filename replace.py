data = open("D:/toolkit_update/ODA/list_temp.txt", "r")
datawrite = open("D:/toolkit_update/ODA/list_temp_co.txt", "w")
for line in data:
        str1="/"
        str2="\\"
        newstr = line.replace(str1, str2)
        datawrite.write("adl_co " + newstr )
      #  print "adl_co " + newstr 

