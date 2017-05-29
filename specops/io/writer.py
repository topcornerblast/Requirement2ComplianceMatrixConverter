import specops.io
import sys
from specops.util import Configuration
import win32com.client as win32

class FileWriter(object):
    '''    
    @author: Steven Hoffman
    @version: 1.0
    @summary: writes directly to a file
    '''
    # Filename and path to write the XML (String)
    _outputFile=None
    # file descriptor 
    _file=0
            
    def __init__(self, outputFile=None):
        '''
        @param outputFile: File where data will be written
        @type outputFile: String
        __init__: constructor
        '''
        self._file=0
        self.setOutputFile(outputFile)

    def setOutputFile(self, outputFile):
        '''
        @param outputFile: output file
        @type outputFile: String 
        setOutputFile: sets the output file where data will be written
        '''
        if self.isOpen():
            self.close()
            
        self._outputFile=outputFile
        
    def isOpen(self):
        '''
        @return: true if the file is open, false otherwise
        @rtype: Boolean
        isOpen: checks to see if a file is currently open
        '''
        return self._file != 0

    def open(self, outputFile=None):
        '''
        @param outputFile: output file
        @type outputFile: String
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        open: opens the file for writing
        '''
        if outputFile != None:
            if outputFile != self._outputFile: # make sure it is not the same file
                self.setOutputFile(outputFile)

        # check to see if the file is already opened
        if self.isOpen():
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('File \''+self._outputFile+'\' already opened\n')
            return;
        
        try:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Opening Output File \''+self._outputFile+'\'\n')
            self._file=open(self._outputFile, specops.io.WRITE_ONLY)
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('No such file or directory: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception opening Writer output File: \'' + self._outputFile + '\'\n')
            raise

    def close(self):
        '''
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        open: flushes the buffer and closes the file
        '''  
        if self.isOpen() == False:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('File \''+self._outputFile+'\' already closed\n')
            return
              
        try:
            # flush data to disk
            self.flush()
            
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Closing Output File \''+self._outputFile+'\'\n')
            self._file.close()
            self._file=0  # use this to check is file is open
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('IOError closing Writer output File: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception closing Writer output File: \'' + self._outputFile + '\'\n')
            raise

    def write(self, theString):
        ''' 
        @param theString: string that will be written to the file
        @type theString: String 
        @raise Exception: General exception causing a failure to write
        write: writes the data to the buffer
        '''
        if self.isOpen() == False:
            sys.stderr.write('Writing Data FAILED.  File not open')
            return
            
        try:
            # write the string to the file immediately
            self._file.write(theString)
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception writing Writer ouptut File: \'' + self._outputFile + '\'\n')
            raise
        
    def flush(self):
        '''
        flush: since all data is immediately written to the file, there is nothing much to do here
        '''
        if Configuration.INSTANCE.getBoolean('DEBUG', False):
            sys.stderr.write('Flushing Data to Output File: \'' + self._outputFile + '\'\n')
            
    def __str__(self) :
            return self.toString()

    def __hash__(self):
        return hash(str(self.__dict__))

    def __eq__(self, other) : 
        return self.toString()==other.toString()

    def toString(self):
        return 'FileWriter(FileName='+self._outputFile+',isOpen='+str(self.isOpen())+')'

class BufferedFileWriter(FileWriter):
    '''    
    @author: Steven Hoffman
    @version: 1.0
    @summary: writes to a file using a buffer to improve performance
    '''
    # Filename and path to write the XML (String)
    _outputFile=None
    # file descriptor 
    _file=0
    # buffer string (String)
    _buffer=''
            
    def __init__(self, outputFile=None):
        '''
        @param outputFile: File where data will be written
        @type outputFile: String
        __init__: constructor
        '''
        self._file=0
        self._buffer=''
        self.setOutputFile(outputFile)

    def setOutputFile(self, outputFile):
        '''
        @param outputFile: output file
        @type outputFile: String 
        setOutputFile: sets the output file where data will be written
        '''
        if self.isOpen():
            self.close()
            
        self._outputFile=outputFile

    def open(self, outputFile=None):
        '''
        @param outputFile: output file
        @type outputFile: String
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        open: opens the file for writing
        '''
        if outputFile != None:
            if outputFile != self._outputFile: # make sure it is not the same file
                self.setOutputFile(outputFile)

        # check to see if the file is already opened
        if self.isOpen():
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('File \''+self._outputFile+'\' already opened\n')
            return;
        
        try:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Opening Output File \''+self._outputFile+'\'\n')
            self._file=open(self._outputFile, specops.io.WRITE_ONLY)
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('No such file or directory: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception opening Writer output File: \'' + self._outputFile + '\'\n')
            raise

    def close(self):
        '''
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        open: flushes the buffer and closes the file
        '''  
        if self.isOpen() == False:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('File \''+self._outputFile+'\' already closed\n')
            return
              
        try:
            # flush data to disk
            self.flush()
            
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Closing Output File \''+self._outputFile+'\'\n')
            self._file.close()
            self._file=0  # use this to check is file is open
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('IOError closing Writer output File: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception closing Writer output File: \'' + self._outputFile + '\'\n')
            raise

    def write(self, theString):
        ''' 
        @param theString: string that will be written to the file
        @type theString: String 
        @raise Exception: General exception causing a failure to write
        write: writes the data to the buffer
        '''         
        if self.isOpen() == False:
            sys.stderr.write('Writing Data FAILED.  File not open')
            return
           
        try:
            self._buffer=self._buffer+theString
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception writing Writer ouptut File: \'' + self._outputFile + '\'\n')
            raise
        
    def flush(self):
        '''
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        flush: writes the data to the file
        '''
        try:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Flushing Data to Output File: \'' + self._outputFile + '\'\n')
            
            if self.isOpen():
                self._file.write(self._buffer)
                self._buffer='' # clear the buffer
            else:
                sys.stderr.write('Flushing Data FAILED.  File not open')
            
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('IOError flushing Writer output File: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception flushing Writer output File: \'' + self._outputFile + '\'\n')
            raise

    def toString(self):
        return 'BufferedFileWriter(FileName='+self._outputFile+',isOpen='+str(self.isOpen())+')'
        
class ComplianceMatrixWriter(FileWriter):
    '''
    Created on Apr 9, 2012
    
    @author: Steven Hoffman
    @version: 1.0
    @summary: writes a Compliance Matrix (CSV file that can be loaded in Excel) 
    '''
    
    # buffer to hold on all requirements that will be written
    _requirementList=None
    # Excel Application
    _excelObject=None
    # Excel Workbook that will the requirements will be written 
    _excelWorkbook=None
    # Excel workbook sheet that the requirements will be written
    _sheet=None
    # all the Cells in the Excel Workbook Sheet
    _cells=None
    
    def __init__(self, outputFile=None):
        '''
        @param outputFile: File where data will be written
        @type outputFile: String
        __init__: constructor
        '''
        super().__init__(outputFile)
        
        if outputFile==None:
            self.setOutputFile(Configuration.INSTANCE.getString('complianceMatrixOutputFile','./complianceMatrix'))
        
        self._requirementList=list()
        
    def isOpen(self):
        '''
        @return: true if the file is open, false otherwise
        @rtype: Boolean
        isOpen: checks to see if a file is currently open
        '''
        return self._excelWorkbook != None
            
    def open(self, outputFile=None):
        '''
        @param outputFile: output file
        @type outputFile: String
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        open: opens the file for writing
        '''
        if outputFile != None:
            if outputFile != self._outputFile: # make sure it is not the same file
                self.setOutputFile(outputFile)

        # check to see if the file is already opened
        if self.isOpen():
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('File \''+self._outputFile+'\' already opened\n')
            return;
        
        try:
            self._excelObject=win32.Dispatch('Excel.Application')
            self._excelWorkbook=self._excelObject.Workbooks.Add(1) 
            self._sheet=self._excelWorkbook.ActiveSheet
            self._cells=self._excelWorkbook.ActiveSheet.Cells
            self._writeHeader()
            
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Opening Output File \''+self._outputFile+'\'\n')
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('No such file or directory: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception opening Writer output File: \'' + self._outputFile + '\'\n')
            raise
            
    def _writeHeader(self):    
        '''
        @raise Exception: General exception writing the individual cells
        _writeHeader: writers the header Row to the spreadsheet
        '''        
        if self.isOpen() == False:
            sys.stderr.write('Writing Header Row Data FAILED.  File not open')
            return
        
        try:
            # writer the header row
            self._cells(1,1).Value='Requirement ID'
            self._sheet.Columns("A:A").ColumnWidth=20.0
            self._cells(1,2).Value='Requirement'
            self._sheet.Columns("B:B").ColumnWidth=50.0
            self._cells(1,3).Value='Meets Requirement (Yes / No / Partial)'
            self._sheet.Columns("C:C").ColumnWidth=35.0
            self._cells(1,4).Value='Comment'
            self._sheet.Columns("D:D").ColumnWidth=50.0
        except Exception:
            print ('Exception: ', sys.exc_info()[0])
            sys.stderr.write('General Exception writing Header Row in Writer output File: \'' + self._outputFile + '\'\n')
            raise
        
    def close(self):
        '''
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        open: flushes the buffer and closes the file
        '''  
              
        try:
            # flush data to disk
            self.flush()
            # format table
            #self._sheet.ListObjects.Add(1,'$A'+str(self._lastRow)+':$C'+str(self._lastRow),None,1).Name = "Table1"
            # save file
            self._excelWorkbook.SaveAs(self._outputFile)
            # close the saved excel spreadsheet
            self._excelObject.Workbooks.Close()
            
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Closing Output File \''+self._outputFile+'\'\n')
            self._file=0  # use this to check is file is open
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('IOError closing Writer output File: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception closing Writer output File: \'' + self._outputFile + '\'\n')
            raise
                
    def write(self, requirement):      
        ''' 
        @param requirement: string that will be written to the file
        @type requirement: String 
        @raise Exception: General exception causing a failure to write
        write: writes the data to the buffer list
        '''  
        if self.isOpen() == False:
            sys.stderr.write('Writing Data FAILED.  File not open')
            return
        
        try:
            # add this to the list to be written
            self._requirementList.append(requirement)
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception writing Compliance Matrix Writer output File: \'' + self._outputFile + '\'\n')
            raise
        
    def flush(self):    
        '''
        @raise IOError: File or directory does not exist or unable to be opened
        @raise Exception: Some other sort of exception  
        flush: writes the data to the file
        '''      
        if self.isOpen() == False:
            sys.stderr.write('Flushing Data FAILED.  File not open')
            return
              
        try:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Flushing Compliance Matrix to Output File: \'' + self._outputFile + '\'\n')
            
            lastRow=2
            for requirement in self._requirementList:
                self._cells(lastRow,1).Value=str(lastRow-1)
                self._cells(lastRow,2).Value=requirement   
                self._cells(lastRow,2).WrapText=True
                lastRow=lastRow+1              
            self._requirementList=list() # clear the buffer      
        except IOError:
            print ('IOError: ', sys.exc_info()[0])    
            sys.stderr.write('IOError flushing Compliance Matrix output File: \'' + self._outputFile + '\'\n')
            raise
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception flushing Compliance Matrix output File: \'' + self._outputFile + '\'\n')
            raise

    def toString(self):
        return 'ComplianceMatrixWriter(FileName='+self._outputFile+',isOpen='+str(self.isOpen())+')'