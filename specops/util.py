import sys
from specops.io.reader import CsvFileReader

class ConfigSingleton:
    '''    
    @author: Steven Hoffman
    @version: 1.0
    @summary: Reads in a configuration properties file so that it can be queried by other classes for configuration information
    '''
    
    # path and filename of the properties file to read in
    _propertiesFile='./config/ComplianceMatrixConverter.properties'
    
    # dictionary to hold all the key/value pairs
    _propertyMap=None
    # indicates if the property file was already read or not
    _readFile=False
    
    def __init__(self, propertiesFile=None):
        '''
        @param propertiesFile: Property File to read in
        @type propertiesFile: String  
        '''
        self._propertyMap=dict()
        self._readFile=False
        
        if propertiesFile is not None:
            self._propertiesFile=propertiesFile
        
    def setPropertiesFile(self, propertiesFile):
        '''
        @param propertiesFile: Property File to read in
        @type propertiesFile: String
        setPropertiesFile: set the filename of the properties file to read in
        '''
        if propertiesFile is not None:
            self._propertiesFile=propertiesFile
            self._readFile=False
        
    def readConfig(self):
        '''
        @raise Exception: error opening config file 
        readConfig: read in the config properties file
        '''
        self._readFile=True
        
        try:
            csvFileReader=CsvFileReader(self._propertiesFile)
            csvFileReader.setDelimiter('=')
            csvFileReader.open()
            
            sys.stderr.write("Reading in Properties File\n")
            
            # get all the columns in the row
            for columns in csvFileReader.readlines():  
                if len(columns) > 1:
                    key=columns[0].strip()
                    value=columns[1].strip()
                    
                    sys.stderr.write(key+"="+value+'\n')
                    self._propertyMap[key]=value
            
            csvFileReader.close()
        except Exception:
            print ('Exception: ', sys.exc_info()[0])    
            sys.stderr.write('General Exception opening config File: \'' + self._propertiesFile + '\'\n')
            raise            
            
    def _getProperty(self, key, defaultValue):
        '''
        @param key: key to lookup in the hashmap
        @type key: String
        @param defaultValue: value to return if key is not in map
        @type defaultValue: Object
        @return: object in hashmap corresponding to the key
        @rtype: Object
        _getProperty: get object out of HashMap corresponding to the key
        '''
        if self._readFile == False:
            self.readConfig()
            
        if key in self._propertyMap:
            return self._propertyMap[key]
        else:
            return defaultValue
            
    def getString(self, key, defaultValue):
        '''
        @param key: key to lookup in the hashmap
        @type key: String
        @param defaultValue: value to return if key is not in map
        @type defaultValue: String
        @return: string in hashmap corresponding to the key
        @rtype: String
        getString: get string out of HashMap corresponding to the key
        '''
        return str(self._getProperty(key, defaultValue))
        
    def getBoolean(self, key, defaultValue):
        '''
        @param key: key to lookup in the hashmap
        @type key: String
        @param defaultValue: value to return if key is not in map
        @type defaultValue: Boolean
        @return: boolean value in hashmap corresponding to the key
        @rtype: Boolean
        getBoolean: get boolean value out of HashMap corresponding to the key
        '''
        return {'true': True, 'false': False}.get(self._getProperty(key, str(defaultValue)).lower())
        
    def getInt(self, key, defaultValue):
        '''
        @param key: key to lookup in the hashmap
        @type key: String
        @param defaultValue: value to return if key is not in map
        @type defaultValue: Int
        @return: integer in hashmap corresponding to the key
        @rtype: Int
        getInt: get integer value out of HashMap corresponding to the key
        '''
        return int(self._getProperty(key, defaultValue))
        
    def getFloat(self, key, defaultValue):
        '''
        @param key: key to lookup in the hashmap
        @type key: String
        @param defaultValue: value to return if key is not in map
        @type defaultValue: Object
        @return: float value in hashmap corresponding to the key
        @rtype: Float
        getFloat: get float value out of HashMap corresponding to the key
        '''
        return float(self._getProperty(key, defaultValue))
                
    def __str__(self):
        return self.toString()
    
    def toString(self):
        return self._propertyMap.__str__()

class Configuration:
    INSTANCE=ConfigSingleton()