import os

# os.mknod("./excel_source/test.sql")
file = open("/Users/gaoxing/pythod_project/generate_insert/excel_source/test1.sql", "a")
file.write("insert into t11 \n")
file.write("insert into t111 \n")
file.close()
