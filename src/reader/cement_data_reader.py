import textract
import re
import pandas as pd

result = textract.process(r"C:\Users\JasSu\Documents\UHD\Sem7\TCOM\project data\report 1.docx",
                          extension='docx', method='wordminer')
print(result)
result_decode = result.decode()

result_sub = re.sub(",", "", result_decode)
result_sub = re.sub("\n", "-", result_sub)
print(result_sub)

mix_strength = re.search("Mix\sStrength:---(\d+)", result_sub)
print(mix_strength.group())

air_temperature = re.search("Air\sTemperature:\s(\d+)", result_sub)
print(air_temperature.group())

concrete_temperature = re.search("Concrete\sTemperature:\s(\d+)", result_sub)
print(concrete_temperature.group())

air_content = re.search("Air\sContent:---(\d+[.]\d+%\s)", result_sub)
print(air_content.group())

# 43-1301-D-653-1--04/29/2020--7--Lab--6.01--28.36--165625--5840--5--U--04/23/2020--ds--
x = re.findall("\d+-\d+-\w-\d+-\d--\d\d[/]\d\d[/]\d\d\d\d--\d+--Lab--\d+[.]\d+--\d+[.]\d+--\d+--\d+--\d+--\w+--\d\d[/]\d\d[/]\d\d\d\d", result_sub)
from collections import defaultdict
data = defaultdict(list)

#Sample #----Test Date----Test Age--Curing Type (F/L)--Average Computed Diameter (in)--Avg.--
# CrossSectional Area (in²)--Breaking Load (lbs)--Rounded Compr. Str. (psi)----Brk. Type----Cap Type----Date to Lab----Lab Tech--
for i in x:
    print(i)
    splits = i.split('--')
    print(len(splits))
    data['MixStrength(psi)']=mix_strength.group().split('---')[1]
    data['AirTemperature(F)']=air_temperature.group().split(': ')[1]
    data['ConcreteTemperature(F)']=concrete_temperature.group().split(': ')[1]
    data['AirContent(%)']=air_content.group().split('---')[1]
    data['Sample'].append(splits[0])
    data['TestDate'].append(splits[1])
    data['TestAge'].append(splits[2])
    data['CuringType'].append(splits[3])
    data['AvgComputedDiameter'].append(splits[4])
    data['CrossSectionalArea'].append(splits[5])
    data['BreakingLoad'].append(splits[6])
    data['RoundedCompressionStrength'].append(splits[7])
    data['BrkType'].append(splits[8])
    data['CapType'].append(splits[9])
    data['DateToLab'].append(splits[10])
    #data['LabTech'].append(splits[11])

df = pd.DataFrame.from_dict(data)
df.to_csv("b.csv")
