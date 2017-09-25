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
A Qt model object for presenting sheets name.
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
* def \_\_init__(self,sheet = None,sheet_wr = None,xmlname_coordinate = None)
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
#### Detailed Description
#### Method Documentation
- - -







