import specops.io
import sys
#import docx
import win32com.client as win32

class FileReader:
    '''    
    @author: Steven Hoffman
    @version: 1.0
    @summary: reads in a file
    '''
    # Filename and path of the file to read in (String)
    _inputFile='NONE'
    # file descriptor
    _file=0

    #c'tor
    def __init__(self, inputFile=None):
        '''
        @param inputFile: File where data will be read
        @type inputFile: String
        __init__: constructor
        '''
        self._file=0
        self.setInputFile(inputFile)
       
    def setInputFile(self, inputFile):
        '''
        @param inputFile: input file
        @type inputFile: String 
        setOutputFile: sets the input file where data will be read
        '''
        # check to make sure a file is not already open
        if self.isOpen():
            sys.stderr.write('Set new Input File while old input file is open.  Closing current input file'+'\n')
            self._file.close()
            self._file=0
        
        self._inputFile=inputFile
       
    def isOpen(self):
        '''
        @return: true if file is open, false otherwise
        @summary: check to see if the file is open for reading
        '''
        return self._file != 0
    
    def open(self, inputFile=None):
        '''
        @param inputFile: File to open for reading
        @type inputFile: String
        @raise IOError: error opening the file or directory
        @raise Exception: general exception 
        @summary: opens the file
        '''
        if inputFile != None:
            self.setInputFile(inputFile)
        
        try:
            self._file=open(self._inputFile, specops.io.READ_ONLY)
        except IOError:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            sys.stderr.write("No such file or directory: '"+self._inputFile+"'\n")
            self.file=0
            return
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            self.file=0
            return
    
    def close(self):
        '''
        @raise Exception: general exception 
        @summary: closes the file
        '''
        try:
            self._file.close()
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return
        
    def readline(self):
        '''
        @summary: reads a single line in the file
        @return: a single line in a file, or None if an exception occurs
        @rtype: String
        @raise Exception: generic exception indicting something went wrong 
        '''
        try:
            return self._file.readline()
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return None
        
    def readlines(self):
        '''
        @summary: reads every line in the file
        @return: list of rows, with each row containing a single line in a file,
                 or None if an exception occurs
        @rtype: list<String>
        @raise Exception: generic exception indicting something went wrong 
        '''
        try:
            return self._file.readlines()
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return None
        
    def __str__(self) :
            return self.toString()

    def __hash__(self):
        return hash(str(self.__dict__))

    def __eq__(self, other) : 
        return self.toString()==other.toString()

    def toString(self):
        return 'FileReader(FileName='+self._inputFile+',isOpen='+str(self.isOpen())+')'


class CsvFileReader(FileReader):
    '''    
    @author: Steven Hoffman
    @version: 1.0
    @summary: Reads in a CSV file and splits it into fields using a delimiter
    '''
    # delimiter used to split lines (String)
    _delimiter=','

    # c'tor
    def __init__(self, inputFile=None):
        '''
        @param inputFile: File where data will be read
        @type inputFile: String
        __init__: constructor
        '''
        super().__init__(inputFile)
        
    def setDelimiter(self, delimiter):
        '''
        @summary: set the delimiter used to split each line of the file
        @param delimiter: string used to split each line of the file
        @type delimiter: String
        '''
        
        self._delimiter=delimiter
    
    def getDelimiter(self):
        '''
        @summary: get the delimiter used to split each line of the file
        @return: the delimiter
        @rtype: String
        '''
        
        return self._delimiter
    
    def readline(self):
        '''
        @summary: reads a single line in the file, with each line being split based on 
                  the delimiter
        @return: list of strings that were split based on the delimiter,
                 or None if an exception occurs
        @rtype: list<String>
        @raise Exception: generic exception indicting something went wrong 
        '''
        
        try:
            line=super().readLine()        
            return line.split(self._delimiter)
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return None
        
    def readlines(self):
        '''
        @summary: reads every line in the file and adds them to a list of rows, with
                  each row containing a list of strings that were split based on 
                  the delimiter
        @return: list of rows, with each row containing a single line in a file.  Each
                 row containing a list of strings that were split based on the delimiter,
                 or None if an exception occurs
        @rtype: list< list<String> >
        @raise Exception: generic exception indicting something went wrong 
        '''
        
        try:
            rows=list()
            for line in super().readlines():
                rows.append(line.split(self._delimiter))
        
            return rows
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return None

    def toString(self):
        return 'CsvFileReader(FileName='+self._inputFile+',isOpen='+str(self.isOpen())+')'
        
        
class WordDocumentFileReader(FileReader):
    '''
    Created on Apr 9, 2012
    
    @author: Steven Hoffman
    @version: 1.0
    @summary: Reads in a Word Document file and reads the body of the file
    '''

    # word document to read from
    _word=None
    # current line number
    _lineNumber=0
    
    # c'tor
    def __init__(self, inputFile=None):
        '''
        @param inputFile: File where data will be read
        @type inputFile: String
        __init__: constructor
        '''
        super().__init__(inputFile)
        self._word=None
        self._lineNumber=0
        
    def isOpen(self):
        '''
        @return: true if file is open, false otherwise
        @summary: check to see if the file is open for reading
        '''
        return self._word != None
    
    def open(self, inputFile=None):
        '''
        @param inputFile: File to open for reading
        @type inputFile: String
        @raise IOError: error opening the file or directory
        @raise Exception: general exception 
        @summary: opens the Word file
        '''
        if inputFile != None:
            self.setInputFile(inputFile)
        
        try:
            self._word=win32.Dispatch('Word.Application') 
            self._word.Documents.Open(self._inputFile)
            # get the total number of paragraphs
            self._lineNumber=0
        except IOError:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            sys.stderr.write("No such file or directory: '"+self._inputFile+"'\n")
            self.file=0
            self._word=None
            self._lineNumber=0
            return
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            self.file=0
            self._word=None
            self._lineNumber=0
            return
        
    def close(self):
        '''
        @raise Exception: general exception 
        @summary: closes the file
        '''
        try:
            # close the open word document
            self._word.Documents[0].Close()
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return
        
    def readline(self):
        '''
        @summary: reads a single line in the Word Document file
        @return: single line from the word file or None if an exception occurs
        @rtype: String
        @raise Exception: generic exception indicting something went wrong 
        '''
        
        try:
            line=None
            if self._lineNumber < self._word.Documents[0].Content.Paragraphs.Count:
                line=str(self._word.Documents[0].Content.Paragraphs[self._lineNumber])
                self._lineNumber=self._lineNumber+1
            return line
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return None
        
    def readlines(self):
        '''
        @summary: reads every line in the file and adds them to a list of rows
        @return: list of rows, with each row containing a single line in a file,
                 or None if an exception occurs
        @rtype: list< list<String> >
        @raise Exception: generic exception indicting something went wrong 
        '''
        
        try:
            rows=list()
            for line in self._word.Documents[0].Content.Paragraphs:
                rows.append( str(line) )
            return rows
        except Exception:
            sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')
            return None

    def toString(self):
        return 'WordDocumentFileReader(FileName='+self._inputFile+',isOpen='+str(self.isOpen())+')'
