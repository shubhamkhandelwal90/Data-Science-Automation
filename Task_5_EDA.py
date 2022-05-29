#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import xlwings as xw

import warnings
warnings.filterwarnings('ignore')

pd.set_option('display.max_columns', 3000)
pd.set_option('display.max_rows', 3000)


df_orders = pd.read_csv('orders_2016-2020_Dataset.csv')
df_review = pd.read_csv('review_dataset.csv')


# The analysis of Reviews given by Customers
def Review_Analysis():
    fig = plt.figure()
    plt.title('Reviews Given By Customers',fontdict={'fontsize':20,'fontweight':15,'color':'green'})
    df_review.stars.value_counts().plot(kind = 'bar')
    plt.ylabel("Number of Reviews", fontdict={'fontsize':20,'fontweight':15,'color':'green'})
    plt.xlabel("Reviews", fontdict={'fontsize':20,'fontweight':15,'color':'green'})
    plt.legend
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Review_Analysis.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Review_Analysis.xlsx")


# The analysis of different payment methods used by the Customers
def Payment_Methods():    
    fig = plt.figure()
    plt.title('Different Payment Methods used by the Customers',fontdict={'fontsize':20,'fontweight':5,'color':'green'})
    df_orders['Payment Method'].dropna().str.split().apply(lambda x: x[0]).value_counts().plot(kind = 'pie',autopct='%.1f%%',
                                figsize = (6,6), textprops={'fontsize': 16, 'fontweight' : 20, 'color' : 'Black'}, startangle=90)
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Payment_Methods.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Payment_Methods.xlsx")
    
# The analysis of Top Consumer States of India
def Top_Consumer_States():    
    fig = plt.figure()
    sns.countplot((df_orders['Billing State']), order=df_orders['Billing State'].value_counts().iloc[:31].index)
    plt.xticks(rotation = 90)
    plt.title('Top Consumer States of India',fontdict={'fontsize':20,'fontweight':5,'color':'green'})
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Top_Consumer_States.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Top_Consumer_States.xlsx")

#The analysis of Top Consumer Cities of India
def Top_Consumer_Cities():
    fig = plt.figure(figsize = (20,20))
    sns.countplot((df_orders['Billing City']), order=df_orders['Billing City'].value_counts().iloc[:31].index)
    plt.xticks(rotation = 90)
    plt.title('Top Consumer Cities of India',fontdict={'fontsize':20,'fontweight':5,'color':'green'})
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Top_Consumer_Cities.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Top_Consumer_Cities.xlsx")
    
# The analysis of Top Selling Product Categories
def Top_Selling_Categories():
    fig = plt.figure(figsize = (10,10))
    sns.countplot((df_review['category']), order=df_review['category'].value_counts().iloc[:31].index)
    plt.xticks(rotation = 90)
    plt.title('Top Selling Product Categories',fontdict={'fontsize':20,'fontweight':5,'color':'green'})
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Top_Selling_Categories.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Top_Selling_Categories.xlsx")   
    
# The analysis of Reviews for All Product Categories
def Review_All_Product_Categories():
    fig = plt.figure(figsize = (10,10))
    plt.title('Reviews for All Product Categories',fontdict={'fontsize':20,'fontweight':15,'color':'green'})
    df_review.groupby('category')['stars'].count().plot(kind = 'bar',figsize=(16, 8))
    plt.ylabel("Number of Reviews", fontdict={'fontsize':20,'fontweight':15,'color':'green'})
    plt.xlabel("Category", fontdict={'fontsize':20,'fontweight':15,'color':'green'})
    plt.legend
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Review_All_Product_Categories.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Review_All_Product_Categories.xlsx") 

# The analysis of Number of Orders Per Month Per Year
def Orders_Per_Month():
    fig = plt.figure()
    plt.title('Number of Orders Per Month', fontsize=18)
    pd.to_datetime(df_orders['Order Date and Time Stamp']).dt.month.value_counts().plot(kind='pie', figsize=(21, 18), autopct='%0.2f%%')
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Orders_Per_Month.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Orders_Per_Month.xlsx") 
    
# The analysis of Reviews for Number of Orders Per Month Per Year
def Reviews_Of_Orders_Per_Month_Per_Year():
    fig = plt.figure()
    plt.title('Reviews for Number of Orders Per Month Per Year', fontsize=18)
    df_review['month'] = pd.to_datetime(df_orders['Order Date and Time Stamp']).dt.month
    df_review.groupby("month")['stars'].value_counts().plot(kind='bar', figsize=(12, 8))
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Reviews_Of_Orders_Per_Month_Per_Year.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Reviews_Of_Orders_Per_Month_Per_Year.xlsx") 
    
# The analysis of Number of Orders Across Parts of a Day
def Order_Across_Part_Of_Day():
    fig = plt.figure()
    plt.title('Number of Orders Across Parts of a Day', fontsize=18)
    Orders = pd.to_datetime(df_orders['Order Date and Time Stamp']).dt.strftime('%H:%M:%S').value_counts().values
    plt.plot(Orders)
    plt.show()
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.name = "excel charts"
    sht.pictures.add(fig, name="Order_Part_day.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Order_Across_Part_Of_Day.xlsx") 

#the Full Report
def Full_Report():
    fig = plt.figure()
    plt.subplot(3, 4, 1)
    plt.title('Review Analysis Given by Constumer', fontsize=20)
    review_count = df_review['stars'].value_counts()
    review_count.plot(kind='bar', label='<2.0 is negative rating')
    plt.ylabel("Number of Reviews", fontsize=20)
    plt.xlabel("Reviews", fontsize=5)
    plt.show
    
    plt.legend
    plt.subplot(3, 4, 2)
    plt.title('Payment method Used By Customers', fontsize=20)
    df_orders['Payment Method'].dropna().str.split().apply(lambda x: x[0]).value_counts().plot(kind='pie', autopct='%0.2f%%',
                                                                                         figsize=(8, 4))
    plt.show
    
    plt.subplot(3, 4, 3)
    plt.title('Top Consumer States of India', fontsize=20)
    df_orders["Billing State"].dropna().value_counts().head().plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    plt.show
    
    plt.subplot(3, 4, 4)
    plt.title('Top Consumer Cities of India', fontsize=20)
    df_orders["Billing City"].dropna().value_counts().head().plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    plt.show
    
    plt.subplot(3, 4, 5)
    plt.title('Top Selling Product Categories', fontsize=20)
    df_review["category"].value_counts().head(10).plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    plt.show
    
    plt.subplot(3, 4, 6)
    plt.title('Reviews for All Product Categories', fontsize=20)
    df_review.groupby('category')['stars'].count().plot(kind='bar', figsize=(40, 50))
    plt.show
    
    plt.subplot(3, 4, 7)
    plt.title('Number of Orders Per Month Per Year', fontsize=20)
    pd.to_datetime(df_orders['Order Date and Time Stamp']).dt.month.value_counts().plot(kind='pie', autopct='%0.2f%%',
                                                                                            shadow=True, figsize=(8, 4))
    plt.show
    
    plt.subplot(3, 4, 8)
    plt.title('Reviews for Number of Orders Per Month Per Year', fontsize=20)
    df_review['month'] = pd.to_datetime(df_orders['Order Date and Time Stamp']).dt.month
    df_review.groupby("month")['stars'].value_counts().plot(kind='bar', figsize=(40, 50))
    plt.show
    
    plt.subplot(3, 4, 9)
    plt.title('Number of Orders Across Parts of a Day', fontsize=20)
    Orders = pd.to_datetime(df_orders['Order Date and Time Stamp']).dt.strftime('%H:%M:%S').value_counts().values
    plt.plot(Orders)
    plt.show
    
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.pictures.add(fig, name="Full_Report.xlsx", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=800)
    wb.save("Full_Report.xlsx") 
   
    
Statement = '''1. Enter 1 to see the analysis of Reviews given by Customers
2. Enter 2 to see the analysis of different payment methods used by the Customers
3. Enter 3 to see the analysis of Top Consumer States of India
4. Enter 4 to see the analysis of Top Consumer Cities of India
5. Enter 5 to see the analysis of Top Selling Product Categories
6. Enter 6 to see the analysis of Reviews for All Product Categories
7. Enter 7 to see the analysis of Number of Orders Per Month Per Year
8. Enter 8 to see the analysis of Reviews for Number of Orders Per Month Per Year
9. Enter 9 to see the analysis of Number of Orders Across Parts of a Day
10. Enter 10 to see the Full Report
11. Enter 11 for exit'''

print(Statement)
while True:
    Input = input('Enter the number to see the task report : ')
    if Input == '1':
        Review_Analysis()
    if Input == '2':
        Payment_Methods()
    if Input == '3':
        Top_Consumer_States()
    if Input == '4':
        Top_Consumer_Cities()
    if Input == '5':
        Top_Selling_Categories()
    if Input == '6':
        Review_All_Product_Categories()
    if Input == '7':
        Orders_Per_Month()
    if Input == '8':
        Reviews_Of_Orders_Per_Month_Per_Year()
    if Input == '9':
        Order_Across_Part_Of_Day()
    if Input == '10':
        Full_Report()
    if Input == '11':
        break


# In[ ]:




