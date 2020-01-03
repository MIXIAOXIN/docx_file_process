import os
import numpy as np
import pandas as pd

in_dir = "/Volumes/LACIE SHARE/7_其他文档/研究生课程/工程案例助教/7/7"
out_dir = "/Volumes/LACIE SHARE/7_其他文档/研究生课程/工程案例助教/7"

out_filename = '第7次作业名单.csv'
out_path = os.path.join(out_dir, out_filename)

for root, dir, files in os.walk(in_dir):
    print('out_path is', out_path)
    student_ids = []
    student_names = []
    student_files = []
    for file in files:
        # print('filename is: ', file)
        if file[0] == '2':
            # print('2XXX filename is: ', file)
            ID = file.split('_')
            NAME = ID[-1].split('.d')
            NAME = NAME[0]

            student_ids.append(ID[0])
            student_names.append(NAME)
            student_files.append(file)

    data_np = [student_ids, student_names, student_files]
    data = pd.DataFrame(data_np, ['ID', 'Name', 'Files'])
    data_t = data.stack().unstack(0)

    data_t.to_csv(out_path)



