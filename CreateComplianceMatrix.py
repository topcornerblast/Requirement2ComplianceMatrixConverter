import sys
import argparse
from specops.io.reader import *
from specops.io.writer import *
from specops.util import Configuration
'''
Created on Apr 9, 2012

@author: Steven Hoffman
@version: 1.0
@requires: Python 3.1 or later
'''

class CreateComplianceMatrix:  
    '''
    Created on Apr 9, 2012
    
    @author: Steven Hoffman
    @version: 1.0
    @summary: Main Class.  Will read in the Word document and create a compliance matrix
    @requires: Python 3.1 or later
    '''
    
    # file to read from
    _inputFile=None
    # file to write to
    _outputFile=None
    # extension of the input file.  Was going to be used to verify that it is a Word Document
    _extensionType=None
        
    def __init__(self, wordFile, outputFile=None):
        '''
        @param wordFile: word file to read in
        @type wordFile: String
        __init__: contructor
        '''
        self._inputFile=wordFile
        
        # get the input filename without extension
        extensionIdx=self._inputFile.rfind('.')
        self._extensionType=self._inputFile[extensionIdx:]
        if outputFile == None:
            self._outputFile=self._inputFile[:extensionIdx]+'.xlsx'
        else:
            self._outputFile=outputFile
        
    # simple tokenizer
    def tokenize(self, sentence, sep):
        '''
        @param sentence: sentence/paragraph to be tokenized
        @type sentence: String
        @param sep: delimiter to use to tokenize into sentences
        @type sep: String
        @return: list of sentence strings
        '''
        containsPeriodException=False
        
        sentences=list()
        
        # make sure the we don't count etc. or i.e. as a complete sentence.  (Add more here if we find any)
        if sep == '.':
            if sentence.find('etc.') != -1:
                sentence=sentence.replace('etc.','etc!@#$!@#$')
                containsPeriodException=True
            if sentence.find('i.e.') != -1:
                sentence=sentence.replace('i.e.','i!@#$!@#$e!@#$!@#$')
                containsPeriodException=True
            if sentence.find('e.g.') != -1:
                sentence=sentence.replace('e.g.','e!@#$!@#$g!@#$!@#$')
                containsPeriodException=True
        start=0
        end=sentence.find(sep)
        while end != -1:
            line=sentence[start:end+1]
            if containsPeriodException == True:
                # change back to a period
                line=line.replace('!@#$!@#$','.')
            sentences.append(line.strip())
            start=end+1
            end=sentence.find(sep,start)
        # make sure the sentence contains something and it isn't just the end of a valid sentence that was added in the loop
        if len(sentence[start:]) != 0:
            sentences.append(sentence[start:])
        return sentences
    
    def generateComplianceMatrix(self):
        '''
        generateComplianceMatrix: Generate the Compliance matrix
        '''
        wordFileReader=None
        complianceMatrixWriter=None
        
        # open the word document
        try:
            wordFileReader=WordDocumentFileReader(self._inputFile)
            wordFileReader.open()
        except IOError:
            return
        except Exception:
            return
        
        try:
            # open the complianceWriter
            complianceMatrixWriter=ComplianceMatrixWriter(self._outputFile)
            complianceMatrixWriter.open()
        except IOError:
            return
        except Exception:
            return
        
        try:
            # get all the lines in the word document
            for line in wordFileReader.readlines():          
                if Configuration.INSTANCE.getBoolean('DEBUG', False):
                    sys.stderr.write('line: '+line+'\n')
        
                # tokenize the line using '.' as a separator
                for sentence in self.tokenize(line,"."):
                    # check if this line is a SHALL statement
                    if sentence.lower().find("shall") != -1:
                        # It is a SHALL statement.  Write it to the complianceMatrixWriter
                        if Configuration.INSTANCE.getBoolean('DEBUG', False):
                            sys.stderr.write('found a SHALL statement\n')
                        complianceMatrixWriter.write(sentence.strip())
        except IOError:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Read Word Document Failed\n')            
        except Exception:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Exception: '+str(sys.exc_info()[0])+'\n')  
            
        
        try:
            wordFileReader.close()
        except IOError:  
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Input File Failed to close')            
        except Exception:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Input File Failed to close')   
        
        try:
            complianceMatrixWriter.close()
        except IOError:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Output File Failed to close')   
        except Exception:
            if Configuration.INSTANCE.getBoolean('DEBUG', False):
                sys.stderr.write('Output File Failed to close')   
    

if __name__ == '__main__':    
    ###
    # MAIN PROGRAM
    #
    
    try:        
        # create a argument parser
        parser=argparse.ArgumentParser(description='Convert a System Specification Word Document into a Excel Compliance Matrix')
        parser.add_argument('--input_file', help='Full Path and Filename of the System Specification Word Document', required=False)
        parser.add_argument('--output_file', help='Full Path and Filename to write the Excel Compliance Matrix', required=False, default=None)
        #parser.add_argument('--help', help='show this help message and exit', nargs=0, required=False, action='store_const')
        
        args=parser.parse_args()
        if args.input_file != None:
            complianceMatrix=CreateComplianceMatrix(args.input_file, args.output_file)
            complianceMatrix.generateComplianceMatrix()
        else:
            print ('--input_file <Word Document> is required')
            parser.print_help()
    except IOError:
        ignore=True
    except KeyboardInterrupt:
        ignore=True
    except Exception:
        print ('Exception: ', sys.exc_info()[0])    
        raise