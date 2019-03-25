import os

result_all = []

def all_path(dirname):
    for maindir, subdir, file_name_list in os.walk(dirname):
        for filename in file_name_list:
            apath = os.path.join(maindir, filename)
            if '.ts' in filename:
                result_all.append(apath)
all_path('./translate')

for i in range(len(result_all)):
    os.system("lrelease.exe {}".format(result_all[i]))