import win32com.client as win32
import os
import numpy as np
import pandas as pd

import string 

class RangeShape1D(Exception):
    """Exception raised for when ranges are not the right shape
    """

    def __init__(self, shape, message="Shape must be 2D for ranges, where each inner row represents rows"):
        self.shape = shape
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        return f'{self.shape} -> {self.message}'
    
class ColShape(Exception):
    """Exception raised for when cols are not the right shape
    """

    def __init__(self, shape, message="Cols must be 1D or the second dim must = 1"):
        self.shape = shape
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        return f'{self.shape} -> {self.message}'
    
class RowShape(Exception):
    """Exception raised for when rows are not the right shape
    """

    def __init__(self, shape, message="rows must be 1D"):
        self.shape = shape
        self.message = message
        super().__init__(self.message)

    def __str__(self):
        return f'{self.shape} -> {self.message}'
    
def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def string_location_to_ij(loc):
    col = []
    row = []
    for c in loc:
        if c in string.ascii_letters:
            col.append(c)
        else:
            row.append(c)

    return int("".join(row)), col2num("".join(col)) 



class work_book:
    def __init__(self, path, live = True, password=None, create = True):
        
        self.path = path
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
           
            self.wb = self.excel.Workbooks.Open(path, Password=password)
            if live: 
                self.excel.Visible = True
            print(f"Found file at {path}")
            #self.establish_tracking_link() # to create the logs page
        except:
            print(f"Couldn't find file at {path}, try using os.getcwd() if it is in the same folder?")
            if create:
                print('But, creating!')
                self.wb = self.excel.Workbooks.Add();
                self.saveAs()
        
    def sheet(self, name):
        print(f"Activating {name}")
        return work_sheet(self.wb, name)
    
    def establish_tracking_link(self):
        # creates the logging page for tracking, you will also need to put the macro in each sheet with the function "insert change tracker"
        if 'logs' in [self.wb.Sheets(i).Name for i in range(1,self.wb.Sheets.Count+1)]:
            self.logs = work_sheet(self.wb, "logs")
        else:
            ws = self.wb.Worksheets.Add()
            ws.Name = "logs"
            self.wb.Worksheets('logs').Visible = 0 # hide the logs sheet
            self.logs = work_sheet(self.wb, "logs") # create one of our objects for it 
            
    def check_for_changes(self):
        changes = self.logs.read_block("A1")
        if changes is None: return None
        change_list = []
        for c in changes.split(','):
            sheet, rng = c.split('-')
            change_list.append((sheet, rng))
        # wipe it so we only keep new changes
        self.logs.write_row("A1", [""])
        return change_list

    def wipe_changes_log(self):
        self.logs.write_row("A1", [""])

            
    def show(self):
        self.excel.Visible = True
        
    def hide(self):
        self.excel.Visible = False
        
    def saveAs(self):
        self.wb.SaveAs(self.path)
       
    def save(self):
        self.wb.Save()
        
    def quit(self):
        self.excel.Application.Quit()
        
class work_sheet:
    def __init__(self, wb, name):
        self.name = name
        self.wb = wb
        self.ws = wb.Worksheets(name)
        
    def insert_tracker(self):
        # this is how to add macros to individual worksheets, only do this if its a new sheet - haven't worked out how to check
        # presence yet
        excelModule = self.wb.VBProject.VBComponents(self.ws.CodeName)
        # Code which tracks changes and puts them into a logs sheet, so that we can check A1 in the logs sheet for changes!
        macro = """
        Public Sub Worksheet_Change(ByVal target As Range)

            Debug.Print "Something changed in cell " & target.Address(0, 0)
            If Worksheets("logs").Range("A1").Value <> "" Then
                Worksheets("logs").Range("A1").Value = Worksheets("logs").Range("A1").Value & ","
            End If
            Worksheets("logs").Range("A1").Value = Worksheets("logs").Range("A1").Value & target.Parent.Name & "-" & target.Address(0, 0)
        End Sub
        """
        excelModule.CodeModule.AddFromString(macro)
        
    def cell(self, i,j, offset = (0,0)):
        return self.ws.Cells(i,j).Offset(offset[0]+1, offset[1]+1)

    def write_range(self, location, rng): 
        '''
        Ranges must come in as 2D, where the inner lists represent the rows, and out represent the cols
        For MVP: assume immediate numpy conversion
        '''
        # later, for efficiency check if it is a pandas series or whaterver
        # but this is a good catch all, which we then convert to list so that
        # the indiviual types are not np types (which causes win32com to fail)
        rng = np.array(rng) 
        shape = rng.shape
        if len(shape) < 2:
            raise RangeShape1D(shape)
        else:
            self.ws.Range(location).Value = tuple(map(tuple, rng.tolist()))
            
    def read_range(self, location):
        return self.ws.Range(location).Value
            
    def write_top_left(self, top_left, offset, shape, block):
        top_left_cell = self.ws.Cells(top_left[0], top_left[1]).Offset(offset[0]+1, offset[1]+1)
        height, width = shape
        bottom_right_cell = self.ws.Cells(top_left[0]+height-1, top_left[1]+width-1).Offset(offset[0]+1, offset[1]+1)
        self.ws.Range(top_left_cell, bottom_right_cell).Value = tuple(map(tuple, block.tolist()))
        
    def write_column(self, top_left, col, offset = (0,0)):
        '''
        Define top in terms of i,j location or colrow string
        '''
        if isinstance(top_left , str):
            top_left = string_location_to_ij(top_left)
        # Columns are meant to be 2D, but in case they come in one d correct that
        col = np.array(col)
        shape = col.shape
        if len(shape) < 2:
            col = np.expand_dims(col,1)
        # now that it is 2d, ensure second dim is 1
        
        shape = col.shape
        if shape[1] != 1:
            raise ColShape(shape)
            
        # ok, validation done
        self.write_top_left(top_left, offset, shape, col)
        
    def write_row(self, top_left, row, offset = (0,0)):
        '''
        Define left in terms of i,j location or colrow string
        '''
        if isinstance(top_left , str):
            top_left = string_location_to_ij(top_left)
        # Columns are meant to be 2D, but in case they come in one d correct that
        row = np.array(row)
        shape = row.shape
        if len(shape) < 2:
            row = np.expand_dims(row,0)
        # now that it is 2d, ensure second dim is 1
        
        shape = row.shape
        if shape[0] != 1:
            raise RowShape(shape)
            
        # ok, validation done
        self.write_top_left(top_left, offset, shape, row)
        
    def write_block(self, top_left, block, offset = (0,0)): 
        '''
        Writes a range, but only marked by the top left of the block
        '''
        if isinstance(top_left , str):
            top_left = string_location_to_ij(top_left)
        # later, for efficiency check if it is a pandas series or whaterver
        # but this is a good catch all, which we then convert to list so that
        # the indiviual types are not np types (which causes win32com to fail)
        block = np.array(block) 
        shape = block.shape
        if len(shape) < 2:
            raise RangeShape1D(shape)
        self.write_top_left(top_left, offset, shape, block)
        
        
    def read_col(self, top, offset= (0,0)):
        return self.get_col(top, offset).Value
    
    def get_col(self, top, offset= (0,0)):
        if isinstance(top , str):
            top = string_location_to_ij(top)
            
        top_cell = self.ws.Cells(top[0], top[1]).Offset(offset[0]+1, offset[1]+1)
        length = 0
        cell = self.ws.Cells(top[0], top[1]).Offset(offset[0]+1, offset[1]+1)
        while(cell.Value != None):
            length = length + 1
            cell = self.ws.Cells(top[0]+length, top[1]).Offset(offset[0]+1, offset[1]+1)
        bottom_cell = self.ws.Cells(top[0]+length-1, top[1]).Offset(offset[0]+1, offset[1]+1)
        length = max(1,length)
        # currently if length is 1, it will just return the value itself.
        return self.ws.Range(top_cell, bottom_cell)
        
    def read_row(self, top, offset= (0,0)):
        return self.get_row(top, offset).Value
    
    def get_row(self, top, offset= (0,0)):
        if isinstance(top , str):
            top = string_location_to_ij(top)
            
        top_cell = self.ws.Cells(top[0], top[1]).Offset(offset[0]+1, offset[1]+1)
        length = 0
        cell = self.ws.Cells(top[0], top[1]).Offset(offset[0]+1, offset[1]+1)
        while(cell.Value != None):
            length = length + 1
            cell = self.ws.Cells(top[0], top[1]+length).Offset(offset[0]+1, offset[1]+1)
        length = max(1, length)
        bottom_cell = self.ws.Cells(top[0], top[1]+length-1).Offset(offset[0]+1, offset[1]+1)

        # currently if length is 1, it will just return the value itself.
        return self.ws.Range(top_cell, bottom_cell)
    

    def read_block(self, top, offset= (0,0), height_init=0):
        '''
        Height init is if you have any indication how many rows there are - because this check is slow as
        To actually find the end - there must be a better way to do it!
        '''
        
        return self.get_block(top, offset, height_init).Value
    
    
    def get_block(self, top, offset= (0,0), height_init=0):
        '''
        Height init is if you have any indication how many rows there are - because this check is slow as
        To actually find the end - there must be a better way to do it!
        '''
        
        if isinstance(top , str):
            top = string_location_to_ij(top)

        top_cell = self.ws.Cells(top[0], top[1]).Offset(offset[0]+1, offset[1]+1)
        height = height_init
        cell = self.ws.Cells(top[0], top[1]).Offset(offset[0]+1, offset[1]+1)
        while(cell.Value != None and cell.Value != ""):
            height =  height + 1
            cell = self.ws.Cells(top[0]+height, top[1]).Offset(offset[0]+1, offset[1]+1)
        width = 0
        cell = self.ws.Cells(top[0], top[1]).Offset(offset[0]+1, offset[1]+1)
        while(cell.Value != None and cell.Value != ""):
            width = width + 1
            cell = self.ws.Cells(top[0], top[1]+width).Offset(offset[0]+1, offset[1]+1)
        width, height = max(1, width), max(1, height)
        print(f"Width - {width}, Height - {height}")
        bottom_cell = self.ws.Cells(top[0]+height-1, top[1]+width-1).Offset(offset[0]+1, offset[1]+1)

        # currently if length is 1, it will just return the value itself.
        return self.ws.Range(top_cell, bottom_cell)



############ helpful functions ###################### 

def as_pandas_df(rng):
    df = pd.DataFrame(list(rng))
    df.columns = df.iloc[0]
    df = df[1:]
    return df
