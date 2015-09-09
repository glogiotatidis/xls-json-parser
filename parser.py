#!/usr/bin/env python

import xlrd
import codecs
from collections import OrderedDict
import simplejson as json

def get_initials(name):
 first=name[0]
 i = name.index(' ')
 last=name[i+1]
 return first+last

doc= xlrd.open_workbook('el.xlsx')
print doc.sheet_names()

sh = doc.sheet_by_index(0)

people_list = []
temp="";
for rownum in range(0, sh.nrows):
 person = OrderedDict()
 row_values = sh.row_values(rownum)
 
 
 for values in row_values:
  if values=='':
   i=row_values.index(values)
   row_values[i]="null"
 
 #if 
 # person['head']= TRUE
 person['name'] = row_values[0]
 person['initials'] = get_initials(person['name'])
 person['party'] = row_values[1]
 person['occupation'] = row_values[2]
 person['twitter'] = row_values[3]	
 person['facebook'] = row_values[4]	
 person['section'] = row_values[5]
 person['head'] = False
 
 if temp!=row_values[5]:
  person['head'] = True
 
 temp = row_values[5]

 people_list.append(person)

#print people_list

j = json.dumps(people_list).encode('utf8')

with codecs.open("data.json", "w", encoding='utf-8') as f:
 f.write(j)




