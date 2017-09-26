# Excel-operator API Documentation
The excel operator is composed by three parts, model to store the data, view to show content and controller to handle all the actions.

## 1.Model
There is six sub-packages in the Model package.
* Workbook
* Workcell
* Worksheet
* CASbook
* CASsheet
* PSbook
* PSsheet

**Folder structure**
```
model
├── __init__.py
├── Workbook.py
├── Workcell.py
├── Worksheet.py
├── CASbook.py
├── CASsheet.py
├── PSbook.py
└── PSsheet.py
```
- - -
### Workbook Package
There is only one class in Workbook package.
* Workbook
- - -
### Workbook Class Reference
The Workbook is a abstract class that used to store the instance created by xlwings and openpyxl.
#### Pubilc Methods
* def \_\_init__(self,workbook = None,app = None)
* def \_\_del__(self)
* def init_book(self,workbook,app)
* def init_model(self)
* def update_model(self)
* def update_sheet_name_model(self)
* def load_sheets(self,sheet_cls,sheets,sheets_wr)
* def load_sheets_name(self,sheets)
* def open(self,path_name)
* def save_as(self,path_name)
* def save(self)
* def recover(self,direction)
* def sheets(self)
* def sheet_name_model(self)
* def workbook(self)
* def workbook_wr(self)
* def workbook_name(self)
* def sheetnames(self)
#### Private Methods
* \_workbook
* \_workbook\_wr
* \_workbook\_name
* \_sheets
* \_sheet\_name\_model
#### Detailed Description
The Workbook provides a series of interface that could be used for both PS file and CAS file. If there are some particular methods only for PS file or CAS file, they will be defined in PSbook class and CASbook class.
#### Method Documentation
##### \_workbook
Store the workbook instance created by openpyxl for reading.
##### \_workbook\_wr
Store the workbook instance created by xlwings for writing.
##### \_workbook\_name
Store the name of the workbook.
##### \_sheets
A list to store all the sheets object.
##### \_sheet\_name\_model
A Qt item model object for presenting sheets name.
##### def \_\_init__(self,workbook = None,app = None)
Create a new Workbook instance with the given **workbook** and **app**.
##### def \_\_del__(self)
Deletes all sheets hold by the workbook instance and release the workbook instance.
##### def init_book(self,workbook,app)
Load workbook by xlwings and openpyxl, the xlwings is used to write and the openpyxl is used to read.
##### def init_model(self)
Initialize the workbook's name and create a new model to hold all the sheet's name.
##### def update_model(self)
Whenever workbook changed, call update_model() to keep the Workbook model update.  
For now, the Workbook only keep a sheet name model.
##### def update_sheet_name_model(self)
Read the sheet's name and fill the model.
##### def load_sheets(self,sheet_cls,sheets,sheets_wr)
Called by the child class.  
**sheet_cls** provide a specified class with some special method that could be applied for the sheet.  
**sheets** is used to read.  
**sheets_wr** is used to write.  
##### def load_sheets_name(self,sheets)
This method is used to load every sheet's name to a list.
##### def open(self,path_name)
Virtual function, only a placeholder that waiting for overwrite by child class.
##### def save_as(self,path_name)
Save the workbook with the given **path_name** by xlwings.
##### def save(self)
Virtual function, only a placeholder that waiting for overwrite by child class.
##### def recover(self,direction)
Virtual function, only a placeholder that waiting for overwrite by child class.
##### def sheets(self)
Property member, provide a interface to access **_sheets**.
##### def sheet_name_model(self)
Property member, provide a interface to access **\_sheet\_name\_model**.
##### def workbook(self)
Property member, provide a interface to access **\_workbook**.
##### def workbook_wr(self)
Property member, provide a interface to access **\_workbook\_wr**.
##### def workbook_name(self)
Property member, provide a interface to access **\_workbook\_name**.
##### def sheetnames(self)
Property member, provide a interface to access **\_workbook.sheetname**.
- - -
### Workcell package
The Workcell package contains four classes.  
* Workcell
* Xmlname
* Header
* Status
- - -
### Workcell Class Reference
#### Public Methods
* def \_\_init__(self,cell = None,sheet_wr = None,header = None,xmlname = None)
* def cell(self)
* def row(self)
* def col(self)
* def column(self)
* def col_letter(self)
* def value(self)
* def value(self,value)
* def header(self)
* def xmlname(self)
* def sheet_wr(self)
#### Private Methods
* \_cell
* \_header
* \_xmlname
* \_sheet_wr
#### Detailed Description
Workcell is a base class that provides interfaces to access cells via xlwings or openpyxl.
#### Method Documentation
##### \_cell
Store the cell instance created by openpyxl.
##### \_header
Store the Header object related to this cell.
##### \_xmlname
Store the Xmlname object related to this cell.
##### \_sheet_wr
Store the sheet which this cell belongs.
##### def \_\_init__(self,cell = None,sheet_wr = None,header = None,xmlname = None)
Create a new Workcell object with the given **cell**, **sheet_wr**, **header**, **xmlname**.
##### def cell(self)
Property member, provide a interface to access **_cell**.
##### def row(self)
Property member, provide a interface to access **_cell.row**.
##### def col(self)
Property member, provide a interface to access **_cell.col_idx**.
##### def column(self)
Property member, provide a interface to access **_cell.col_idx**.
##### def col_letter(self)
Property member, provide a interface to access **_cell.column**.
##### def value(self)
Property member, provide a interface to read **_cell.value**.  
It is completed by openpyxl.
##### def value(self,value)
Property member, provide a interface to write **_sheet_wr.api.Range(self.row,self.column).value**.  
It is completed by xlwings.
##### def header(self)
Property member, provide a interface to access **_header**.
##### def xmlname(self)
Property member, provide a interface to access **_xmlname**.
##### def sheet_wr(self)
Property member, provide a interface to access **_sheet_wr**.
- - -
### Xmlname Class Reference
Inherits **Workcell**.
#### Public Methods
* def \_\_init__(self,cell = None,sheet_wr = None)
* def get_item_by_header(self,header)
#### Private Methods
Inherits Workcell.
#### Detailed Description
Xmlname is a child class of Workcell. It is used to store the xmlname cells.
#### Method Documentation
##### def \_\_init__(self,cell = None,sheet_wr = None)
Create a new Xmlname object with given **cell** and **sheet_wr**.
##### get_item_by_header(self,header)
Return the cell which is located at the cross of the row of xmlanme and the column of header. 
- - -
### Header Class Reference
Inherits **Workcell**.
#### Public Methods
* def \_\_init__(self,cell = None,sheet_wr = None)
* def get_item_by_xmlname(self,xmlname)
#### Private Methods
Inherits Workcell.
#### Detailed Description
Header is a child class of Workcell. It is used to store the header cells.
#### Method Documentation
##### def \_\_init__(self,cell = None,sheet_wr = None)
Create a new Header object with given **cell** and **sheet_wr**.
##### get_item_by_xmlname(self,xmlname)
Return the cell which is located at the cross of the row of xmlanme and the column of header. 
- - -
### Status Class Reference
Inherits **Workcell**.
#### Public Methods
* def \_\_init__(self,cell = None,sheet_wr = None)
* def get_item_by_xmlname(self,xmlname)
#### Private Methods
Inherits Workcell.
#### Detailed Description
Status is a child class of Workcell. It is used to store the status cells.
#### Method Documentation
##### def \_\_init__(self,cell = None,sheet_wr = None)
Create a new Header object with given **cell** and **sheet_wr**.
##### get_item_by_xmlname(self,xmlname)
Return the cell which is located at the cross of the row of xmlanme and the column of Status.  
- - -
### Worksheet package
The Worksheet package contains four classes.  
* Worksheet
* QPreviewItem
* QComparisonItem
* QHeaderItem
- - -
### Worksheet Class Reference
#### Public Methods
* def \_\_init__(self,sheet = None,sheet_wr = None)
* def \_\_del__(self)
* def init_sheet(self)
* def init_model(self) 
* def update_model(self)
* def search_by_value(self,value)
* def search_header_by_value(self,value)
* def search_xmlname_by_value(self,value)
* def xml_names(self)
* def xml_names_value(self)
* def headers(self)
* def headers_value(self)
* def select_all_headers(self)
* def unselect_all_headers(self)
* def cell(self,row,col)
* def cell_value(self,row,col)
* def cell_wr(self,row,col)
* def cell_wr_value(self,row,col)
* def xmlname(self)
* def preview_model(self)
* def header_model(self)
* def header_list(self)
* def xml_name_model(self)
* def extended_preview_model(self)
* def checked_headers(self)
* def worksheet(self)
* def rows(self)
* def cols(self)
* def max_row(self)
* def min_row(self)
* def max_col(self)
* def min_col(self)
#### Private Methods
* \_worksheet
* \_worksheet_wr
* \_xmlname
* \_header_model
#### Detailed Description
Worksheet is a abstract class to define the general methods of worksheet that could be applied on both PS file and CAS file. If there are some particular methods only for PS file or CAS file, they will be defined in PSsheet class and CASsheet class.
#### Method Documentation
##### \_worksheet
Store the worksheet object created by openpyxl for reading.
##### \_worksheet_wr
Store the worksheet object created by xlwings for writing.
##### \_xmlname
Store the 'xmlname' cell object. If there is no 'xmlname', the value will be 'None'.
##### \_header_model
A Qt item model for storing the all the headers.
##### def \_\_init__(self,sheet = None,sheet_wr = None)
Create a new Worksheet object with the given **sheet** and **sheet_wr**.
##### def \_\_del__(self)
Release all the **_worksheet** and **_worksheet_wr**.
##### def init_sheet(self)
Initialize the **_xmlname**.
##### def init_model(self) 
Create a Qt item model to store all the headers.
##### def update_model(self)
Reload the **_header_model** and keep the model update.
##### def search_by_value(self,value)
Search and return the Workcell object by the value of targe cell.
##### def search_header_by_value(self,value)
Search and return the Header object by the value of target cell.
##### def search_xmlname_by_value(self,value)
Search and return the Xmlname object by the value of target cell.
##### def xml_names(self)
Return a list of Xmlname object which contains all the xmlnames in this worksheet.
##### def xml_names_value(self)
Return a list of xmlname string of all the xmlnames in this worksheet.
##### def headers(self)
Return a list of Header object which contains all the headers in this worksheet.
##### def headers_value(self)
Return a list of header string of all the headers in this worksheet.
##### def select_all_headers(self)
Set the states of all the items of **_header_model** to **Qt.Checked**.
##### def unselect_all_headers(self)
Set the states of all the items of **_header_model** to **Qt.Unchecked**.
##### def cell(self,row,col)
Return the specified cell with the given **row** and **col** via openpyxl.
##### def cell_value(self,row,col)
Retrun the value of specified cell with the given **row** and **col** via openpyxl.
##### def cell_wr(self,row,col)
Return the specified cell with the given **row** and **col** via xlwings.
##### def cell_wr_value(self,row,col)
Return the value of specified cell with the given **row** and **col** via xlwings.
##### def xmlname(self)
Property member, provide a interface to access the **_xmlname**.
##### def preview_model(self)
Property member, provide a interface to access the **_preview_model**.
##### def header_model(self)
Property member, provide a interface to access the **_header_model**.
##### def header_list(self)
Property member.
Return a list of header string of all the headers in this worksheet.
If there is not a 'xmlname' cell, return a blank list instead.
##### def xml_name_model(self)
Property member, provide a interface to access the **_xml_name_model**.
##### def extended_preview_model(self)
Property member, provide a interface to access the **_extended_preview_model**.
##### def checked_headers(self)
Return a list of all the checked item in **_header_model**.
##### def worksheet(self)
Property member, provide a interface to access **_worksheet**.
##### def rows(self)
Property member, provide a interface to access **_worksheet.rows**.
##### def cols(self)
Property member, provide a interface to access **_worksheet.columns**.
##### def max_row(self)
Property member, provide a interface to access **_worksheet.max_row**.
##### def min_row(self)
Property member, provide a interface to access **_worksheet.min_row**.
##### def max_col(self)
Property member, provide a interface to access **_worksheet.max_column**.
##### def min_col(self)
Property member, provide a interface to access **_worksheet.min_column**.
- - -
### QPreviewItem Class Reference
Inherits **QStandardItem**.
#### Public Methods
* def \_\_init__(self,cell)
* def cell(self)
* def value(self)
* def row(self)
* def col(self)
* def col_letter(self)
#### Private Methods
* \_cell
#### Detailed Description
QPreviewItem inherits from QStandardItem.
#### Method Documentation
##### \_cell
Store the Workcell object.
##### def __init__(self,cell)
Create a new QPreviewItem object with the given **cell**.
##### def cell(self)
Property member, provide a interface to access the **_cell**.
##### def value(self)
Property member, provide a interface to access the **_cell.value**.
##### def row(self)
Property member, provide a interface to access the **_cell.row**.
##### def col(self)
Property member, provide a interface to access the **_cell.col**.
##### def col_letter(self)
Property member, provide a interface to access the **_cell.col_letter**.
- - -
### QComparisonItem Class Reference
Inherits **QStandardItem**.
#### Public Methods
* def \_\_init__(self,cell)
* def cell(self)
* def value(self)
* def row(self)
* def col(self)
* def col_letter(self)
#### Private Methods
* \_cell
#### Detailed Description
QComparisonItem inherits from QStandardItem.
#### Method Documentation
##### \_cell
Store the Workcell object.
##### def __init__(self,cell)
Create a new QComparisonItem object with the given **cell**.
##### def cell(self)
Property member, provide a interface to access the **_cell**.
##### def value(self)
Property member, provide a interface to access the **_cell.value**.
##### def row(self)
Property member, provide a interface to access the **_cell.row**.
##### def col(self)
Property member, provide a interface to access the **_cell.col**.
##### def col_letter(self)
Property member, provide a interface to access the **_cell.col_letter**.
- - -
### QHeaderItem Class Reference
Inherits **QStandardItem**.
#### Public Methods
* def \_\_init__(self,cell)
* def get_item_by_xmlname(self,xmlanme)
* def cell(self)
* def value(self)
* def row(self)
* def col(self)
* def col_letter(self)
#### Private Methods
* \_cell
#### Detailed Description
QHeaderItem inherits from QStandardItem.
#### Method Documentation
##### \_cell
Store the Workcell object.
##### def __init__(self,cell)
Create a new QHeaderItem object with the given **cell**.
##### def get_item_by_xmlname(self,xmlname)
Return the cell object which is located at the cross of the row of xmlname and the column of **_cell**.
##### def cell(self)
Property member, provide a interface to access the **_cell**.
##### def value(self)
Property member, provide a interface to access the **_cell.value**.
##### def row(self)
Property member, provide a interface to access the **_cell.row**.
##### def col(self)
Property member, provide a interface to access the **_cell.col**.
##### def col_letter(self)
Property member, provide a interface to access the **_cell.col_letter**.
- - -
### CASbook package
The CASbook package contains one class.  
* CASbook
- - -
### CASbook Class Reference
Inherits **Workbook**.
#### Public Methods
* def \_\_init__(self,file_name = None,app = None)
* def init_cas_book(self)
#### Private Methods
Inherits from Workbook class.
#### Detailed Description
The only difference between CASbook and PSbook is about the **load_sheets**.  
**load_sheets** is used to create specified Worksheet objects for each sheet and load them into a list.  
For CASbook, the specified Worksheet object should be CASsheet.
#### Method Documentation
##### def \_\_init__(self,file_name = None,app = None)
Create a new CASbook objects with the given **file_name** and **app**.
##### def init_cas_book(self)
Create CASsheet objects for each worksheet and load them into a list by **load_sheets**.
- - -
### CASsheet package
The CASsheet package contains one class.  
* CASsheet
- - -
### CASsheet Class Reference
Inherits **Worksheet**.
#### Public Methods
*  def \_\_init__(self,sheet = None,sheet_wr = None)
*  def init_cas_sheet(self)
*  def cell(self,row,col)
#### Private Methods
* \_subject_matter
* \_container_name
#### Detailed Description
While initializing CASsheet object, it will search for two headers, 'Subject Matter Functional Area' and 'Container Name Technical Specification'.
#### Method Documentation
##### \_subject_matter
Store the **QHeaderItem** object as the searching result of cell 'Subject Matter Functional Area'.
##### \_container_name
Store the **QHeaderItem** object as the searching result of cell 'Container Name Technical Specification'.
#####  def \_\_init__(self,sheet = None,sheet_wr = None)
Create a new CASsheet object with the given **sheet** and **sheet_wr**.
#####  def init_cas_sheet(self)
If there is a 'xmlname' cell in this worksheet, search for 'Subject Matter Functional Area' and 'Container Name Technical Specification' and store the result.
#####  def cell(self,row,col)
Return the cell object via xlwings.
- - -
### PSbook package
The PSbook package contains one class.  
* PSbook
- - -
### PSbook Class Reference
Inherits **Workbook**.
#### Public Methods
* def \_\_init__(self,file_name = None,app = None)
* def init_ps_book(self)
#### Private Methods
Inherits from Workbook class.
#### Detailed Description
The only difference between CASbook and PSbook is about the **load_sheets**.  
**load_sheets** is used to create specified Worksheet objects for each sheet and load them into a list.  
For PSbook, the specified Worksheet object should be PSsheet.
#### Method Documentation
##### def \_\_init__(self,file_name = None,app = None)
Create a new PSbook objects with the given **file_name** and **app**.
##### def init_cas_book(self)
Create PSsheet objects for each worksheet and load them into a list by **load_sheets**.
- - -
### PSsheet package
The PSsheet package contains one class.  
* PSsheet
- - -
### PSsheet Class Reference
Inherits **Worksheet**.
#### Public Methods
* def \_\_init__(self,sheet = None,sheet_wr = None)
* def \_\_del__(self)
* def init_ps_sheet(self)
* def init_ps_model(self)
* def update_model(self)
* def status(self)
* def cell(self,row,col)
* def auto_fit(self,cols)
* def add_row(self,start_pos,offset,orientation)
* def delete_row(self,start_pos,offset)
* def lock_row(self,row,status)
* def lock_sheet(self)
* def unlock_sheet(self)
* def unlock_all_cells(self)
* def extended_preview_model(self)
* def extended_preview_model_list(self)
* def preview_model(self)
#### Private Methods
* \_status
* \_subject_matter
* \_container_name
* \_preview_model
* \_preview_model_list
* \_extended_preview_model
* \_extended_preview_model_list
#### Detailed Description
PSsheet class provides more interfaces than Worksheet.The class hold two data model, **_preview_model** for the four columns preview window and **_extended_preview_model** for the full content preview mode.It also allow user to append, delete and lock rows in the worksheet.
#### Method Documentation
##### \_status
Store the **QHeaderItem** object as the searching result of cell 'Status(POR,INIT,PREV)'.
##### \_subject_matter
Store the **QHeaderItem** object as the searching result of cell 'Subject Matter Functional Area'.
##### \_container_name
Store the **QHeaderItem** object as the searching result of cell 'Container Name Technical Specification'.
##### \_preview_model
Store the data model for the content of four columns preview window.
##### \_preview_model_list
Convert **\_preview_model** to list for transmitting between multiprocessing.
##### \_extended_preview_model
Store the data model for the full content of preview window.
##### \_extended_preview_model_list
##### def \_\_init__(self,sheet = None,sheet_wr = None)
Create a new PSsheet object with the given **sheet** and **sheet_wr**.  
Initialize the **_preview_model**, **_preview_model_list**, **_extended_preview_model** and **_extended_preview_model_list**.  
Search for all the headers and constract the **_preview_model**.
##### def \_\_del__(self)
Release the **_preview_model** and **_extended_preview_model**.
##### def init_ps_sheet(self)
Search for headers that 
##### def init_ps_model(self)
Constract the data model for four columns preview window.
##### def update_model(self)
Re-initialize the **_preview_model** and keep the content update.
##### def status(self)
Return a list of Status object which contains all the status items in this worksheet.
##### def cell(self,row,col)
Return the cell object for reading via openpyxl.
##### def auto_fit(self,cols)
Adjust the columns width automatically with the given **cols**.
**cols** should be a collection of the columns you want to adjust.
##### def add_row(self,start_pos,offset,orientation)
Insert several rows below the specified row with the given **start_pos**, **offset**, **orientation**.  
**start_pos** indicates the position of the row that you want to insert below.  
**offset** represents the number of rows you want to insert.  
**orientation** represents the direction of insertion.
##### def delete_row(self,start_pos,offset)
Delete several rows from the specified row with the given **start_pos**, **offset**.  
**start_pos** indicates the start position of the row that you want to delete.  
**offset** represents the number of rows you want to delete.
##### def lock_row(self,row,status)
Set the row's protection mode with the given **row** and **status**.  
**row** represents the row you want to handle.  
**status** represents the target status you want to set. For example, **True** stands for locked and **False** stands for unlocked.
##### def lock_sheet(self)
Set the worksheet's protection mode to locked.  
##### def unlock_sheet(self)
Set the worksheet's protection mode to unlocked.  
##### def unlock_all_cells(self)
Set the protection mode of all the cells in this worksheet to unlocked.
##### def extended_preview_model(self)
Work through the whole worksheet and constrcut a data model for the full content preview window.  
The construction of data model would only be executed for the first time. The following calls would return the result of the first call.
##### def extended_preview_model_list(self)
Convert **_extended_preview_model** to string list to adapter the multiproccessing communication on Windows.
##### def preview_model(self)
Property member, provide a interface to access the **_preview_model_list**.
- - -











