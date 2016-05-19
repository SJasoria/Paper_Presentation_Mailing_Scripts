import xlrd
import string
import xlwt


def lEmail(filename):
   book=xlrd.open_workbook(filename)
   sheet=book.sheet_by_index(0)
   no_rows=sheet.nrows
   dict_email={}
   result={}
   #loop to list all outastation participants
   for row in range(no_rows):
      college= sheet.cell(row,3).value
      if "pilani" in string.lower(college):
        continue
      name=sheet.cell(row,1).value
      email=sheet.cell(row, 2).value
      name=name.lower()
      name=name.title()
      dict_email[name]=email
   #to remove duplicates
   for key, value in dict_email.items():
      if value not in result.values():
        result[key]=value
   print result
   return result

def lEmail4(filename):
   book=xlrd.open_workbook(filename)
   sheet=book.sheet_by_index(0)
   no_rows=sheet.nrows
   details=[]
   for row in range(no_rows):
     ele=[]
     name=sheet.cell(row,0).value
     email=sheet.cell(row,1).value
     name=name.lower()
     name=name.title()
     ele.append(name)
     ele.append(email)
     details.append(ele)
   return details


def lEmail5(filename):
   book=xlrd.open_workbook(filename)
   sheet=book.sheet_by_index(0)
   no_rows=sheet.nrows
   details=[]
   for row in range(no_rows):
     ele=[]
     name=sheet.cell(row,0).value
     email=sheet.cell(row,1).value
     name=name.lower()
     name=name.title()
     ele.append(name)
     ele.append(email)
     details.append(ele)
   print "Before removing redundancies, number of elements "+str(len(details))
   #result=set(details)
   result=details[0]   
   for i in range(len(details)):     #to remove redundancies
      flag=0
      for j in range(len(result)):
          if  result[j][1]==details[i][1]:
            flag=1
            continue
      if flag==0:
         result.append(details[i])      
   return result



def lEmail2(filename):
     book=xlrd.open_workbook(filename)
     sheet=book.sheet_by_index(0)
     no_rows=sheet.nrows
     start=input('Start at? ')
     dict_email={}
     result={}
     for row in range(start,start+50):
          name=sheet.cell(row,0).value
          email=sheet.cell(row,1).value
          name=name.lower()
          name=name.title()
          dict_email[name]=email
          print 'Row: ', row, 'Name: ', name, 'Email: ',email
     print dict_email
     print 'Length of dict_email: ', len(dict_email)
     for key, value in dict_email.items():
          if value not in result.values():
             result[key]=value
     print result
     return result

def correct_case(x):
     x=x.lower()
     x=x.title()
     return x

def lEmail3(filename):
   book=xlrd.open_workbook(filename)
   sheet=book.sheet_by_index(0)
   no_rows=sheet.nrows
   details=[]
   for row in range(no_rows):                              
     ele=[]
     title=sheet.cell(row,1).value
     title=correct_case(title)
     title=title.replace(".","")
     title=title.replace("'","")
     author_name=sheet.cell(row,5).value
     author_name=correct_case(author_name)
     coauthor_name=sheet.cell(row,9).value
     coauthor_name=correct_case(coauthor_name)
     ref_no=sheet.cell(row,3).value
     time=sheet.cell(row,13).value
     time=time.replace('a','')
     room=sheet.cell(row,14).value
     room=int(room)

     print "Datatype of room " , type(room)    #test
     room=int(room)
     ele.append(title)
     ele.append(sheet.cell(row,2).value)
     ele.append(author_name)
     ele.append(sheet.cell(row,7).value)  
     ele.append(coauthor_name)      
     ele.append(sheet.cell(row,11).value)
     ele.append(ref_no)
     ele.append(sheet.cell(row,4).value)
     ele.append(time)  #time
     ele.append(room)  #room

     print ele     #test
     details.append(ele)

   print details[0]
   return details 


def lEmail6(filename):
   book=xlrd.open_workbook(filename)
   sheet=book.sheet_by_index(0)
   no_rows=sheet.nrows
   details=[]
   for row in range(no_rows):
     ele=[]
     category=sheet.cell(row,0).value
     position=sheet.cell(row,1).value
     title=sheet.cell(row,2).value
     title=correct_case(title)
     author_name=sheet.cell(row,3).value
     author_name=correct_case(author_name)
     author_email=sheet.cell(row,5).value
     coauthor_name=sheet.cell(row,6).value
     coauthor_email=correct_case(coauthor_name)
     coauthor_email=sheet.cell(row,8).value
     ele.append(category)
     ele.append(position)
     ele.append(title)
     ele.append(author_name)
     ele.append(author_email)
     ele.append(coauthor_name)
     ele.append(coauthor_email)
     #print ele #test
     details.append(ele)
   #print details
   return details





   

def mail_log(sent, not_sent):
   book=xlwt.Workbook()
   sheet1=book.add_sheet('Sent')
   row=0
   for key, value in sent.items():
      sheet1.write(row, 0, key)
      sheet1.write(row, 1, value)
      row+=1
   sheet2=book.add_sheet('Not Sent')
   row=0
   for key, value in not_sent.items():
      sheet2.write(row, 0, key)
      sheet2.write(row, 1, value)
      row+=1       
   book.save("mail_log.xls")