#!/usr/bin/env python3

import pandas as pd
import openpyxl
import math
import sys
from datetime import date

sale_filename = 'SalesSummaries/SalesSummary_'+str(date.today())+'.xlsx'
SALES_DAYS = 30
SINCE_LAST_SHOP = 7
SHOP_SIZE = 'Small Shop' # 'Normal' is the other option
PREVIOUS_TO_BUY_FILE_NAME = 'To_Buy/2025-10-13_toBuy.xlsx'

class Product:
    def __init__(self, name, case_size, days_to_have, on_hand, location):
        self.name = name
        self.case_size = int(24 if math.isnan(case_size) else case_size)
        self.days_to_have = int(10 if math.isnan(days_to_have) else days_to_have)
        self.on_hand = on_hand
        self.num_sold=0
        self.case_full_amount = 0
        self.location = location
        
        
def process_row(row):
    p1 = Product(row['Product Name'],row['Case size'], row[SHOP_SIZE], 0, '')
    products[p1.name] = p1
    

products = {}

if len(sys.argv) > 1:
    sale_filename = sys.argv[1]

file_path2 = 'SalesSummaries/products.xlsx'
df2 = pd.read_excel(file_path2, sheet_name='Sheet1')
#print(df2)
df2.apply(process_row, axis=1)

file_path3 = PREVIOUS_TO_BUY_FILE_NAME
df3 = pd.read_excel(file_path3)
for index,row in df3.iterrows():
    curr_product = row['Product Name']
    if curr_product in products:
        products[curr_product].on_hand=row['On Hand']+row['To Buy']
        products[curr_product].location=row['Garage Location']

file_path = sale_filename
df = pd.read_excel(file_path)
#print (df)
for index,row in df.iterrows():
    curr_product = row['Product Name']
    if curr_product in products:
        sold = products[curr_product].num_sold+row['Total Vend'] #  do we need to do this?  do products show up more than once?
        sold_per_day = sold / SALES_DAYS
        sold_since_last_shop = sold_per_day
        products[curr_product].num_sold=sold_since_last_shop
        products[curr_product].on_hand=products[curr_product].on_hand-(sold_since_last_shop / products[curr_product].case_size )
        case_full_amount = (sold_per_day * products[curr_product].days_to_have) / products[curr_product].case_size
        products[curr_product].case_full_amount = case_full_amount
    elif 'Monster Java' in curr_product :
        print('javas found: ', row['Total Vend'])
        sold = products['Monster Java'].num_sold+row['Total Vend']
        sold_per_day = sold / SALES_DAYS
        sold_since_last_shop = sold_per_day
        products['Monster Java'].num_sold=sold_since_last_shop
        case_full_amount = (sold_per_day * products['Monster Java'].days_to_have) / products['Monster Java'].case_size
        products['Monster Java'].case_full_amount = case_full_amount
    elif 'Muffin' in curr_product :
        sold = products['Muffin mixed'].num_sold+row['Total Vend']
        sold_per_day = sold / SALES_DAYS
        sold_since_last_shop = sold_per_day
        products['Muffin mixed'].num_sold=sold_since_last_shop
        case_full_amount = (sold_per_day * products['Muffin mixed'].days_to_have) / products['Muffin mixed'].case_size
        products['Muffin mixed'].case_full_amount = case_full_amount 
    elif 'Grandma' in curr_product :
        sold = products['Grandma\'s Cookie Mixed'].num_sold+row['Total Vend']
        sold_per_day = sold / SALES_DAYS
        sold_since_last_shop = sold_per_day
        products['Grandma\'s Cookie Mixed'].num_sold=sold_since_last_shop
        case_full_amount = (sold_per_day * products['Grandma\'s Cookie Mixed'].days_to_have) / products['Grandma\'s Cookie Mixed'].case_size
        products['Grandma\'s Cookie Mixed'].case_full_amount = case_full_amount 
    else:
        print ('Not finding : ', curr_product)
        new_product = Product(curr_product,24,2,0,'Sams Liberty')
        new_product.num_sold = row['Total Vend']
        products[curr_product] = new_product
        
# fix on hand for Monsters and Muffins
    
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = 'Product Inventory and To Buy'

data = [['Product Name','To Buy','Avg Daily Sold','Case Size','Days to Have','On Hand','Full Stock Amount','To Buy Raw','Garage Location']]
counter = 2
for p in products:
    #if products[p].location == 'Sams Liberty' or products[p].location == 'Vistar':
       # print (products[p].name,' ',products[p].num_sold, ' ', products[p].case_full_amount,  ' ', products[p].case_size)
    #if products[p].num_sold!=0:
    to_buy_raw = '=G'+str(counter) + '-F'+str(counter)  # create a formula in the spreadsheet to adjust there if I want
    to_buy_final = '=IF(H'+str(counter)+'<0,0,IF(AND(H'+str(counter)+'>0.25,H'+str(counter)+'<=1),1,ROUND(H'+str(counter)+',0)))'
    cases_full = '=((C'+str(counter)+'/'+str(SALES_DAYS)+')*E'+str(counter)+')/D'+str(counter)
    # =IF(H125<0,0,IF(AND(H125>0.25,H125<=1),1,ROUND(H125,0)))
    # print (to_buy_final)
    data.append([products[p].name, to_buy_final, products[p].num_sold, products[p].case_size, products[p].days_to_have, products[p].on_hand, cases_full, to_buy_raw, products[p].location])
    counter = counter + 1
    # if products[p].case_full_amount > products[p].on_hand:
    #    print('gonna need more ',products[p].name, ' - short by ', products[p].case_full_amount - products[p].on_hand, ' cases for ', products[p].days_to_have )
for row in data:
    sheet.append(row)
   
newFileName = 'To_Buy/' + str(date.today()) + '_toBuy.xlsx'
workbook.save(newFileName)

print('Excel file ', newFileName, ' created successfully.')
