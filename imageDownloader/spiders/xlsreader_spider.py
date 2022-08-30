from importlib.resources import path
import scrapy
 
# import openpyxl module
import openpyxl

class XlsReaderSpider(scrapy.Spider):
  name="xlsreader"
  # Give the location of the file
  path = "G:\\mrCod3r\\projects\\15.scrapper\\imageDownloader\\foods.xlsx"
  
  # To open the workbook
  # workbook object is created
  wb_obj = openpyxl.load_workbook(path)
  
  # Get workbook active sheet object
  # from the active attribute
  sheet_obj = wb_obj.active
  max_col = sheet_obj.max_column

  
  # Loop will print all columns name
  for i in range(1, max_col + 1):
      cell_obj = sheet_obj.cell(row = 1, column = i)
      print('================================')
      if(cell_obj.value == 'banner' or cell_obj.value == 'thumbnail'):
        print(cell_obj)
      # print(type(sheet_obj._cells))
  
