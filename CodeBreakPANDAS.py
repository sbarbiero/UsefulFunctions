import string
import csv
import pandas


t = str.maketrans(dict.fromkeys('!"#$%&\'()*+,-./:;<=>?@[\]^`{|}~', " "))
# make filename a list, and add a loop to do batch in one shot
filename = 'allcodeSOI.txt'

def code_to_list(codefile):
  with open(codefile) as f:
      temp_str_raw = [line.lstrip('\n').rstrip('\n').lower().translate(t).split() for line in f]
      
  code_list = [item for sublist in temp_str_raw for item in sublist]
  
  return code_list
  
flat_list = code_to_list(filename)
    
print (pandas.DataFrame(flat_list)[0].value_counts())

pandas.DataFrame(flat_list)[0].value_counts().to_csv('outputPandasSOI.csv', sep = ',', encoding ='utf-8')
