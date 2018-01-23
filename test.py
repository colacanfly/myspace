#!/usr/bin/python2.7
#!-*- coding:UTF-8 -*-

import xlrd
import os
import MySQLdb

def funtion(root):
    for i in os.listdir(root):
        if os.path.isfile(os.path.join(root,i)):
            if os.path.splitext(i)[1] == ".xls":
                data = xlrd.open_workbook(os.path.join(root,i))
                count=len(data.sheets())
                conn=MySQLdb.connect('localhost','root','123456','testdb')
                cusor=conn.cursor()
                for j in range(count):
                    table = data.sheets()[j]
                    nrows = table.nrows
                    for x in range(2, nrows):
                        row = table.row_values(x)
                        sql="insert into goods(type,number,name,price,link) values(\""+table.name.encode('utf-8')+"\",\""+str(row[0])+"\",\""+row[1].encode('utf-8')+"\",\""+str(row[2])+"\",\""+row[3].encode('utf-8')+"\")"

                        print sql
                        cusor.execute("set names utf8")
                        cusor.execute(sql)
                        conn.commit()
                conn.close()
        else:
            funtion(os.path.join(root,i))
    return;
root='/home/myroot'
funtion(root)

