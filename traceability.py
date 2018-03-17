import os
from string import Template
import subprocess
import hashlib
import six
import logging
import argparse
import csv
import enum

from lxml import etree

from generator import TraceabilityGenerator
from utils import tRequirementLink, tRequirementValue, tLinkType

def buildParser():
    ''' Builds command line argument parser'''
    
    parser = argparse.ArgumentParser(description='Utility for generating requirements traceability documentation')
    
    # action arguments
    parser.add_argument('--TRACE', 
        help='Generate traceability documentation', 
        action='store_true',
        default=False)
    parser.add_argument('--EXPORT', 
        help='Export requirement modules to excel workbooks', 
        action='store_true',
        default=False)
    parser.add_argument('--REPORT', 
        help='Generates text report of all unlinked requirements', 
        action='store_true',
        default=False)
    parser.add_argument('--JENKINS',
        help='Generates XML summary table compatible with Jenkins XML plugin',
        action='store_true',
        default=False)
    
    # configuration arguments
    parser.add_argument('-configFile', 
        help='Parse configuration from configuration file. Supports YAML and JSON', 
        metavar='filename', 
        action='store', 
        nargs=1)
    parser.add_argument('--checkSrcLinks',
        help='Check all requirements are linked to source code',
        action='store_true',
        default=False)
    parser.add_argument('--checkTestLinks',
        help='Check all requirements are linked to test cases',
        action='store_true',
        default=False)
    parser.add_argument('-loggingLevel',
        help='Configure logging level. Supports DEBUG, INFO, WARNING, and ERROR',
        metavar='LEVEL',
        action='store',
        nargs=1)
    
    # input arguments
    parser.add_argument('-modules',
        help='List of requirements module names.',
        metavar='module',
        action='store',
        default=[],
        nargs='+')
    parser.add_argument('-srcDirs',
        help='List of source code directories',
        metavar='directory',
        action='store',
        default=[],
        nargs='+')
    parser.add_argument('-tideDirs',
        help='List of TIDE directories for parsing test code links',
        metavar='directory',
        action='store',
        default=[],
        nargs='+')
    parser.add_argument('-rpyFiles',
        help='List of IBM Rhapsody project files',
        metavar='filename',
        action='store',
        default=[],
        nargs='+')
    
    # output arguments
    parser.add_argument('-outfile',
        help='Specify base name of generated files', 
        metavar='filename', 
        action='store', 
        nargs=1)
    parser.add_argument('-outputDir',
        help='Specify output directory. Defaults to current working directory',
        metavar='directory',
        action='store',
        nargs=1)
        
    # IBM DOORS arguments
    parser.add_argument('-doorsUsr',
        help='DOORS username. Required if export specified.',
        metavar='username',
        action='store',
        nargs=1)
    parser.add_argument('-doorsPwd',
        help='Password for specified DOORS user. Required if export specified.',
        metavar='password',
        action='store',
        nargs=1)
    parser.add_argument('-doorsServer',
        help='DOORS server for connecting to DOORS database. Required if export specified',
        metavar='server',
        action='store',
        nargs=1)
    parser.add_argument('-doorsExe',
        help='Path to DOORS executable. Required if export specified.',
        metavar='exe_path',
        action='store',
        nargs=1)
    parser.add_argument('-doorsView',
        help='DOORS view to use when exporting modules. Required if export specified.',
        metavar='view_name',
        action='store',
        nargs=1)
    
    return parser
    
def addReqLink(reqName, link, reqMap):
    ''' Add requirement link to requirement in module map'''
    
    logger = logging.getLogger(__name__)
    
    isValidReqName = False
    
    # search for requirement name in each module
    for module in six.itervalues(reqMap):
        if (reqName in module):
            isValidReqName = True
            
            if (module[reqName].reqLinks is None):
                module[reqName].reqLinks = [link]
            else:
                # check if link are mapped to requirement
                if (link not in module[reqName].reqLinks):
                    module[reqName].reqLinks.append(link)
            
            break
    
    # check if requirement name found in a module
    if (not isValidReqName):
        logger.warn('Requirement(%s) not found in any requirement modules' % (reqName))

def getFilename(refId, doxygenDirectory):
    ''' get filename of file linked to a requirement by reference id'''
    
    # get base reference id
    baseRefId, _ = refId.rsplit('_', 1)
    # get filename of file containing reference id
    baseRefFile = os.path.join(doxygenDirectory, baseRefId + '.xml')
    
    # parse requirement links file generated by doxygen
    refTree = None
    try:
        refTree = etree.parse(baseRefFile)
    except:
        raise Exception('Failed to read/parse XML file:\n\t%s' % (baseRefFile))
    
    # find XML node with file location
    locationNode = refTree.find('//*[@id=\'%s\']/location' % (refId))
    
    if (locationNode is None):
        raise Exception('Missing file location for reference %s in:\n\t%s' % (refId, baseRefFile))
    
    # parse filename and line number
    filename = locationNode.get('file', None)
    lineNum = locationNode.get('line', None)
    
    if (filename is None):
        raise Exception('Missing file attribute for location element for %s in %s\n\t' % (refId, baseRefFile))
    elif (lineNum is None):
        raise Exception('Missing line attribute for location element for %s in %s\n\t' % (refId, baseRefFile))
    
    lineNum = int(lineNum)
    
    return filename, lineNum

def parseDoxygenReqLinks(srcDir, outputDir, reqType, reqMap):
    ''' Parse requirements linked to test code using doxygen'''
    
    logger = logging.getLogger(__name__)
    
    # make sub-directories for doxygen output
    if (not os.path.exists(outputDir)):
        os.makedirs(outputDir)
    
    # update doxyfile template with the specified source directory and output directory
    cwd = os.path.abspath(os.getcwd())
    doxyTemplateFile = os.path.join(cwd, 'template.doxyfile')
    doxyFile = os.path.join(outputDir, 'project.doxyfile')
    
    if (True != os.path.isfile(doxyTemplateFile)):
        logger.error('Doxygen template does not exist. Expected:\n\t%s' % (doxyTemplateFile))
        return -1
     
    with open(doxyTemplateFile, 'r') as infile:
        template = Template(infile.read())
        with open(doxyFile, 'w') as outfile:
            outfile.write(template.safe_substitute(src_dir=srcDir, output_dir=outputDir))

    try:
        # use doxygen to generate XML documentation
        proc = subprocess.Popen(['doxygen', doxyFile], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = proc.communicate()
        
        # only log non-error output for debugging purposes
        logger.debug(stdout)
        
        # check for any doxygen warnings/errors in STDERR
        if stderr is not None:
            for line in stderr.decode("utf-8") .split('\n'):
                if ('warning:' in line):
                    logger.warn('Doxygen warning(%s) while processing:\n\t%s' % (line[:-1], srcDir))
                elif ('error:' in line):
                    logger.error('Doxygen warning(%s) while processing:\n\t%s' % (line[:-1], srcDir))
    except:
        logger.error('Failed to generate doxygen documentation for:\n\t%s' % (srcDir), exc_info=True)
        return -1
        
    doxygenDir = os.path.join(outputDir, 'xml')
    
    reqXml = os.path.join(doxygenDir, 'REQUIREMENT_LINK.xml')
    
    # requirements XML file only generated if at least one link found
    if (True != os.path.isfile(reqXml)):
        # no requirements found
        logger.debug('No requirements XML generated for:\n\t%s' % (srcDir))
        return 0
    
    # parse requirement links from generated XML documentation
    return parseDoxygenXmlReqLinks(doxygenDir, reqType, reqMap)

def parseSourceReqLinks(srcDir, outputDir, reqMap):
    ''' Parse requirements linked to source code using doxygen'''
    
    logger = logging.getLogger(__name__)
    
    srcDir = os.path.expanduser(srcDir)
    srcDir = os.path.expandvars(srcDir)
    
    logger.info('Parsing source code requirement links for:\n\t%s' % (srcDir))
    
    outputDir = os.path.join(outputDir, 'doxygen', 'src', hashlib.sha1(srcDir.encode('utf-8')).hexdigest())
    
    return parseDoxygenReqLinks(srcDir, outputDir, tLinkType.LINK_TYPE__SRC, reqMap)

def parseRhapsodyModelLinks(rpyFile, reqMap):
    ''' Parse requirement links in model objects in a IBM Rhapsody Project'''
    
    logger = logging.getLogger(__name__)
    
    rpyFile = os.path.expanduser(rpyFile)
    rpyFile = os.path.expandvars(rpyFile)
    
    logger.info('Parsing source links in IBM Rhapsody project:' % (rpyFile))
    return 0

def parseTideTestLinks(tideDir, outputDir, reqMap):
    ''' Parse requirement links in test code in TIDE projects'''
    
    logger = logging.getLogger(__name__)
    
    tideDir = os.path.expanduser(tideDir)
    tideDir = os.path.expandvars(tideDir)
    
    logger.info('Searching for TIDE projects in:\n\t%s' % (tideDir))
    
    if (True != os.path.isdir(tideDir)):
        logger.error('Invalid TIDE directory:\n\t%s' % (tideDir))
        return -1
   
    # search for all TIDE projects in directory
    for root, dirs, _ in os.walk(tideDir):
        for subDir in dirs:
            projectFile = os.path.join(root, subDir, '.project')
            
            # check if project file exists
            # ignore all directories with no project file
            if (True != os.path.isfile(projectFile)):
                continue
            
            projectDir = os.path.join(root, subDir)
            parseTideProjecLinks(projectDir, outputDir, reqMap)
            
    return 0

def parseTideProjecLinks(tideDir, outputDir, reqMap):
    ''' Parse requirement links from TIDE projects within the
    specified directory'''
    
    logger.info('Parsing test code requirement links for:\n\t%s' % (tideDir))

    tideDir = os.path.expanduser(tideDir)
    tideDir = os.path.expandvars(tideDir)

    testDir = os.path.join(tideDir, 'tests')
    
    # check if project has tests
    # ignore projects with no tests
    if (True != os.path.isdir(testDir)):
        return 0
    
    outputDir = os.path.join(outputDir, 'doxygen', 'test', hashlib.sha1(tideDir.encode('utf-8')).hexdigest())
    
    return parseDoxygenReqLinks(tideDir, outputDir, tLinkType.LINK_TYPE__TEST, reqMap)

def parseDoxygenXmlReqLinks(doxygenDirectory, reqType, reqMap):
    ''' Parse requirements linkage XML document generated by doxygen'''
    
    logger = logging.getLogger(__name__)
    
    doxygenDirectory = os.path.expanduser(doxygenDirectory)
    doxygenDirectory = os.path.expandvars(doxygenDirectory)
    
    reqXml = os.path.join(doxygenDirectory, 'REQUIREMENT_LINK.xml')
        
    # parse requirement links 
    reqTree = None
    try:
        reqTree = etree.parse(reqXml)
    except:
        logger.error('Failed to read/parse requirement links XML file:\n\t%s' % (reqXml), exc_info=True)
        return -1

    # iterate through all requirement link lists
    for listNode in reqTree.getroot().findall('.//variablelist'):
        childIterator = iter(listNode)
        
        # list contains pairs of a varlistentry and listitem
        
        listEntryNode = six.next(childIterator, None)
        
        while (listEntryNode is not None):
            if ('varlistentry' == listEntryNode.tag):
                # parse entry type and entry value
                textIter = listEntryNode.itertext()
                # ignore reference type
                six.advance_iterator(textIter)
                refTag = " ".join(map(str.strip, textIter))
            
                # get ref element for entry
                referenceNode = listEntryNode.find('.//ref')
                
                if (referenceNode is None):
                    logger.error('Invalid XML format for req.xml, missing ref element for varlistentry element in file:\n\t%s' % (reqXml))
            
                # get entry's reference id and reference kind
                refId = referenceNode.get('refid', None)
                refKind = referenceNode.get('kindref', None)

                if (refId is None):
                    logger.error('Invalid XML format for req.xml, missing refid attribute for ref element in file:\n\t%s' % (reqXml))
                    return -1
                elif (refKind is None):
                    logger.error('Invalid XML format for req.xml, missing kindref attribute for ref element in file:\n\t%s' % (reqXml))
                    return -1
                elif ('member' != refKind):
                    logger.warn('Unsupported kindref attribute type: %s' % (refKind))
                    listEntryNode = six.next(childIterator, None)
                    continue
                
                # get filename of reference
                filename, lineNum = getFilename(refId, doxygenDirectory)
                
                # get listitem paired with listentry
                listItemNode = six.next(childIterator, None)
                
                if ((listItemNode is None) or ('listitem' != listItemNode.tag)):
                    logger.error('Invalid XML format for req.xml, missing listitem element after varlistentry for %s in file\n\t%s' % (refId, reqXml))
                    return -1
                
                # parse name of each linked requirement
                for itemNode in listItemNode:
                    if ((itemNode.text is not None) and ('' != itemNode.text)):
                        reqName = itemNode.text.strip()
                        addReqLink(reqName, tRequirementLink(reqType, refTag, filename, lineNum), reqMap)
        
            listEntryNode = six.next(childIterator, None)
            
    return 0

def exportDoorsModules(modules, doorsUsr, doorsPwd, doorsServer, doorsView, doorsExe, outputDir='.'):
    ''' Exports a list of IBM DOORS modules to CSV files'''
    
    logger = logging.getLogger(__name__)
    
    outputDir = os.path.expanduser(outputDir)
    outputDir = os.path.expandvars(outputDir)
    
    # make sub-directories for exported DOORS modules
    if (not os.path.exists(outputDir)):
        os.makedirs(outputDir)
    
    # update template DXL script for exporting
    cwd = os.path.abspath(os.getcwd())
    exportTemplate = os.path.join(cwd, 'export_template.dxl')
    exportFile = os.path.join(outputDir, 'export.dxl')
    
    if (True != os.path.isfile(exportTemplate)):
        logger.error('Export DXL script template does not exist. Expected:\n\t%s' % (exportTemplate))
        return -1
        
    moduleExport = '{'
    for module in modules:
        logger.info('Exporting %s to:\n\t%s' % (module, os.path.join(outputDir, module + '.csv')))
        moduleExport += '\'' + module + '\', '
    
    moduleExport = moduleExport.rsplit(',', 1)[0] + '}'
    
    moduleExport = moduleExport.replace('\\', '/')
    view = doorsView.replace('\\', '/')
    output_dir = outputDir.replace('\\', '/')
    
    with open(exportTemplate, 'r') as infile:
        template = Template(infile.read())
        with open(exportFile, 'w') as outfile:
            outfile.write(template.safe_substitute(modules=moduleExport, output_dir=outputDir, view=view))
            
    # use IBM DOORS to export modules to CSV
    doorsExe = os.path.join(doorsExe, 'DOORS.exe')
    if (True != os.path.isfile(doorsExe)):
        logger.error('Invalid path to DOORS executable:\n\t%s' % (doorsExe))
        return -1
    
    try:
        # build call to DOORS exe using user, password, and script
        cmd = [doorsExe, "-W", "-data", doorsServer, "-u", doorsUsr, "-P", doorsPwd, "-b", exportFile]
    
        # export modules
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = proc.communicate()
        
        # only log STDOUT for debug purposes
        logger.debug(stdout)
        
        # check for any errors in processing command
        if ((stderr is not None) and (b'' != stderr)):
            logger.error('Failed to export modules')
            return -1
    except:
        logger.error('Failed to export modules', exc_info=True)
        return -1
        
    return 0

def buildReqMap(modules, outputDir):
    ''' Build a requirement map based on the specified requirement modules'''
    
    logger = logging.getLogger(__name__)
    
    if ((modules is None) or (0 == len(args.modules))):
        logger.error('No requirements modules specified')
        return -1, None
    
    # initialize requirement map
    reqMap = {}
    
    # parse requirements from each module to build initial requirements map
    for moduleName in modules:
        moduleFile = os.path.join(outputDir, moduleName + '.csv')
        
        # parse module requirements
        errCode, moduleMap = parseReqCsv(moduleName, moduleFile)
        
        # add module to requirements map
        if (0 != errCode):
            return -1, None
        
        reqMap[moduleName] = moduleMap
            
    return 0, reqMap

def parseReqCsv(moduleName, moduleFile):
    ''' Parse requirements CSV file'''
    
    logger = logging.getLogger(__name__)
    
    moduleMap = {}
    
    try:
        # parser JSON file
        with open(moduleFile, 'r') as csvfile:
            reader = csv.DictReader(csvfile)
            
            class tReqCsvColHeader(enum.Enum):
                COL_HEADER__REQUIREMENT_NAME = 'ID'
                COL_HEADER__REQUIREMENT_TEXT = 'SW Requirements'
            
            # verify expected column headers in CSV file
            for col in tReqCsvColHeader:
                if (col.value not in reader.fieldnames):
                    logger.error('Expected column(\'%s\') in module CSV:\n\t%s' % (col.value, moduleFile))
                    return -1, None
            
            # parse requirements
            for row in reader:
                # parse requirement details
                reqName = row[tReqCsvColHeader.COL_HEADER__REQUIREMENT_NAME.value]
                reqText = row[tReqCsvColHeader.COL_HEADER__REQUIREMENT_TEXT.value]
                
                # build requirement value based on requirement text, requirement links
                req = tRequirementValue(reqText, [])
                
                # check if requirement already in module map
                if (reqName in moduleMap):
                    logger.warn('Duplicate requirement names(%s) found in module(%s)' % (reqName, moduleName))
                
                # add requirement to module
                moduleMap[reqName] = req
    except:
        logger.error('Unable to parse requirements module:\n\t%s' % (moduleFile), exc_info=True)
        return -1, None
        
    return 0, moduleMap

def parseJsonConfig(configFile, args):
    ''' Parse JSON configuration file for additional command line arguments'''
    
    import json
    
    logger = logging.getLogger(__name__)
    
    try:
        with open(configFile, 'r') as jsonFile:
            jsonConfig = json.load(jsonFile)
    except:
        logger.error('Invalid JSON format for configuration file:\n\t%s' % (configFile), exc_info=True)
        return -1
    
    # update args with parsed jsonConfig
    for key, value in six.iteritems(jsonConfig):
        logger.debug('json config item: (name: %s, type: %s)' % (key, type(value)))
        
        # check if JSON argument is supported
        if key in args:
            # check if argument is a boolean
            if (True == isinstance(vars(args)[key], bool)):
                if (True == isinstance(value, str)):
                    value = value.lower()
                    if ('true' == value):
                        vars(args)[key] = True
                    elif ('false' == value):
                        vars(args)[key] = False
                    else:
                        logger.error('Expected boolean type for argument %s in configuration file:\n\t%s' % (key, configFile))
                        return - 1
                else:
                    logger.error('Expected boolean type for argument %s in configuration file:\n\t%s' % (key, configFile))
                    return - 1
            # check if argument is a list
            elif (True == isinstance(vars(args)[key], list)):
                if (True == isinstance(value , list)):
                    for valueItem in value:
                        if (True == isinstance(valueItem, str)):
                            vars(args)[key].append(valueItem)
                        else:
                            logger.error('Expected string list for argument %s in configuration file:\n\t%s' % (key, configFile))
                            return - 1
                else:
                    logger.error('Expected string list for argument %s in configuration file:\n\t%s' % (key, configFile))
                    return - 1
            # default to using string for other arguments
            else:
                if (True == isinstance(key, str)):
                    vars(args)[key] = value
                else:
                    logger.error('Expected string type for argument %s in configuration file:\n\t%s' % (key, configFile))
                    return - 1
        else:
            logger.warn('Unsupported argument(%s) in configuration file:\n\t%s' % (key, configFile))
     
    return 0

def parseYamlConfig(configFile, args):
    ''' Parse YAML configuration for additional command line arguments'''
    
    import yaml 
    
    logger = logging.getLogger(__name__)
    
    try:
        with open("config.yml", 'r') as ymlfile:
            yamlConfig = yaml.load(ymlfile)
    except:
        logger.error('Invalid YAML format for configuration file:\n\t%s' % (configFile), exc_info=True)
        return -1
    
    # update args with parsed yamlConfig
    for key, value in six.iteritems(yamlConfig):
        logger.debug('YAML config item: (name: %s, type: %s)' % (key, type(value)))
        
        # check if YAML argument is supported
        if key in args:
            # check if argument is a boolean
            if (True == isinstance(vars(args)[key], bool)):
                if (True == isinstance(value, str)):
                    value = value.lower()
                    if ('true' == value):
                        vars(args)[key] = True
                    elif ('false' == value):
                        vars(args)[key] = False
                    else:
                        logger.error('Expected boolean type for argument %s in configuration file:\n\t%s' % (key, configFile))
                        return - 1
                else:
                    logger.error('Expected boolean type for argument %s in configuration file:\n\t%s' % (key, configFile))
                    return - 1
            # check if argument is a list
            elif (True == isinstance(vars(args)[key], list)):
                if (True == isinstance(value , list)):
                    for valueItem in value:
                        if (True == isinstance(valueItem, str)):
                            vars(args)[key].append(valueItem)
                        else:
                            logger.error('AExpected string list for argument %s in configuration file:\n\t%s' % (key, configFile))
                            return - 1
                else:
                    logger.error('BExpected string list for argument %s in configuration file:\n\t%s' % (key, configFile))
                    return - 1
            # default to using string for other arguments
            else:
                if (True == isinstance(key, str)):
                    vars(args)[key] = value
                else:
                    logger.error('Expected string type for argument %s in configuration file:\n\t%s' % (key, configFile))
                    return - 1
        else:
            logger.warn('Unsupported argument(%s) in configuration file:\n\t%s' % (key, configFile))
    
    return 0

def parseConfig(configFile, args):
    ''' Parse configuration file for additional command line arguments'''
    
    logger = logging.getLogger(__name__)
    
    logger.debug('Parsing configuration file:\n\t%s' % (configFile))
    
    if (True != os.path.isfile(configFile)):
        logger.error('Invalid configuration file:\n\t%s' % (configFile))
        return -1
    
    _, ext = os.path.splitext(configFile)
    
    if ('.yaml' == ext.lower()):
        return parseYamlConfig(configFile, args)
    elif ('.json' == ext.lower()):
        return parseJsonConfig(configFile, args)
    else:
        logger.error('Unsupported configuration file type:\n\t%s' % (configFile))
        return -1
    
    return 0

def configureOutput(args):
    ''' Configure output for items generated by script'''
    
    logger = logging.getLogger(__name__)
    
    if (args.outputDir is None):
        args.outputDir = os.getcwd()
    else:
        args.outputDir = os.path.expanduser(args.outputDir)
        args.outputDir = os.path.expandvars(args.outputDir)

        try:
            # make directory path if doesn't already exist
            if (not os.path.exists(args.outputDir)):
                os.makedirs(args.outputDir)
        except:
            logger.error('Failed to create output directory:\n\t%s' % (args.outputDir))
            return -1
        
    return 0

def configureLogger(args):
    ''' Configure logger based on parsed arguments '''

    logger = logging.getLogger(__name__)
    
    # configure logger if specified
    LOG_FILENAME = os.path.join(args.outputDir, 'log.txt')
    fh = logging.FileHandler(LOG_FILENAME)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    if (None != args.loggingLevel):
        try:
            logger.setLevel(logging.getLevelName(args.loggingLevel))
            fh.setLevel(logging.getLevelName(args.loggingLevel))
            logger.debug('Setting logging level to %s' % (logger.getEffectiveLevel()))
        except:
            logger.error('Unsupported logging level:\n\t%s' % (args.loggingLevel))
            return -1
    else:
        logger = logging.getLogger(__name__)
        logger.setLevel(logging.INFO)
        fh.setLevel(logging.INFO)
        logger.debug('No logging level specified. Defaulting to INFO.')
    
    logger.addHandler(fh)
    
    return 0

def parseConfigFile(args):
    ''' Parse configuration file for additional command-line arguments'''
    # parse configuration file(s) if provided
    for filename in args.configFile:
        if (0 != parseConfig(filename, args)):
            return -1
        
    return 0

def validateArgs(args):
    ''' Validate command line arguments'''
    
    logger = logging.getLogger(__name__)
    
    # validate at least one action selected
    if ((True != args.EXPORT) and 
        (True != args.TRACE) and 
        (True != args.JENKINS) and 
        (True != args.REPORT)):
        logger.error('At least one action must be specified (EXPORT, TRACE, JENKINS, or REPORT')
        return -1
    
    # validate DOORS arguments
    if (True == args.EXPORT):
        if (args.doorsUsr is None):
            logger.error('DOORS user must be specified if EXPORT is selected')
            return -1
        elif (args.doorsPwd is None):
            logger.error('DOORS password must be specified if EXPORT is selected')
            return -1
        elif (args.doorsServer is None):
            logger.error('DOORS server must be specified if EXPORT is selected')
            return -1
        elif (args.doorsExe is None):
            logger.error('DOORS executable must be specified if EXPORT is selected')
            return -1
        elif (args.doorsView is None):
            logger.error('DOORS view must be specified if EXPORT is selected')
            return -1
    
    # validate output arguments
    if ((args.outfile is None) or ('' == args.outfile)):
        logger.debug('No outfile basename specified. Defaulting to \'traceability\'')
        args.outfile = 'traceability'
    else:
        logger.debug('Using outfile basename: %s' % (args.outfile))

    return 0

def generateReport(reqMap, args):
    ''' Generate report of all unmapped requirements'''
    
    logger = logging.getLogger(__name__)

    reportFile = os.path.join(args.outputDir, args.outfile + '_report.txt')
    logger.info('Generating missing requirements report:\n\t%s' % (reportFile))
    
    with open(reportFile, 'w') as f:
        # check each module
        for moduleName, module in six.iteritems(reqMap):
            # check each requirement in module
            for req, reqValue in six.iteritems(module):
                if (True == args.checkSrcLinks):
                    # check source code links
                    hasSrcLink = False
                    for link in reqValue.reqLinks:
                        if (tLinkType.LINK_TYPE__SRC == link.linkType):
                            hasSrcLink = True
                            break
                    if (True != hasSrcLink):
                        f.write('[WARNING] %s::%s has no source code link\n' % (moduleName, req))
                        
                if (True == args.checkTestLinks):
                    # check test links
                    hasTestLink = False
                    for link in reqValue.reqLinks:
                        if (tLinkType.LINK_TYPE__TEST == link.linkType):
                            hasTestLink = True
                            break
                    if (True != hasTestLink):
                        f.write('[WARNING] %s::%s has no test link\n' % (moduleName, req))

def generateJenkinsSummary(reqMap, args):
    ''' Generate summary table of requirements links for Jenkins Summary Display plugin'''
    
    logger = logging.getLogger(__name__)

    summaryFile = os.path.join(args.outputDir, args.outfile + '_summary.xml')
    logger.info('Generating requirements linkages summary:\n\t%s' % (summaryFile))
    
    # generate XML table for the Jenkins Summary Display plugin
    root = etree.Element('root')
    table = etree.SubElement(root, 'table')
    
    # add column headers
    titleRow = etree.SubElement(table, 'tr')
    etree.SubElement(titleRow, 'td', attrib={'fontattribute':'bold', 'align':'center'}).text = 'Module Name'
    
    etree.SubElement(titleRow, 'td', attrib={'fontattribute':'bold', 'align':'center'}).text = '# Reqs'
    
    if (True == args.checkSrcLinks):
        etree.SubElement(titleRow, 'td', attrib={'fontattribute':'bold', 'align':'center'}).text = 'Source Links (%)'
    
    if (True == args.checkTestLinks):
        etree.SubElement(titleRow, 'td', attrib={'fontattribute':'bold', 'align':'center'}).text = 'Test Links (%)'
    
    # add row per module
    for moduleName, module in six.iteritems(reqMap):
        moduleRow = etree.SubElement(table, 'tr')
        
        etree.SubElement(moduleRow, 'td', attrib={'fontattribute':'bold', 'align':'center'}).text = moduleName
        
        numReqs = 0
        isModuleFullyLinked = True
        
        if (True == args.checkSrcLinks):
            numSrcLinks = 0
        
        if (True == args.checkTestLinks):
            numTestLinks = 0
        
        # calculate requirement link summaries
        for reqValue in six.itervalues(module):
            numReqs += 1
        
            if (True == args.checkSrcLinks):
                # check for a source code link
                for link in reqValue.reqLinks:
                    if (tLinkType.LINK_TYPE__SRC == link.linkType):
                        numSrcLinks += 1
                        break
            
            if (True == args.checkTestLinks):
                # check for a test code link
                for link in reqValue.reqLinks:
                    if (tLinkType.LINK_TYPE__TEST == link.linkType):
                        numTestLinks += 1
                        break
        
        etree.SubElement(moduleRow, 'td', attrib={'align':'center'}).text = str(numReqs)
        
        if (True == args.checkSrcLinks):
            reqPercent = 0
            if (0 != numReqs):
                # calculate percent met
                reqPercent = 100*(numSrcLinks / numReqs)
                
                # check if all requirements met
                if (numSrcLinks != numReqs):
                    isModuleFullyLinked = False
            etree.SubElement(moduleRow, 'td', attrib={'align':'center'}).text = '%.2f' % (reqPercent)
        
        if (True == args.checkTestLinks):
            reqPercent = 0
            if (0 != numReqs):
                # calculate percent met
                reqPercent = 100*(numTestLinks / numReqs)
                
                # check if all requirements met
                if (numTestLinks != numReqs):
                    isModuleFullyLinked = False
            etree.SubElement(moduleRow, 'td', attrib={'align':'center'}).text = '%.2f' % (reqPercent)
        
        # if all requirements in module are not met, highlight module row with red
        if (True != isModuleFullyLinked):
            for child in moduleRow:
                child.set('bgcolor', 'red')
                
    with open(summaryFile, 'wb') as f:
        f.write(etree.tostring(root, pretty_print=True))
        
if '__main__' == __name__:
    logger = logging.getLogger(__name__)
    
    parser = buildParser()
    args = parser.parse_args()
    
    # check if configuration file provided for additional arguments
    if (args.configFile is None):
        args.configFile = []
        
    errCode = parseConfigFile(args)
    if (0 != errCode):
        exit(errCode)
    
    # configure utility output
    errCode = configureOutput(args)
    if (0 != errCode):
        exit(errCode)
        
    # configure logger
    errCode = configureLogger(args)
    if (0 != errCode):
        exit(errCode)
    
    # any log statements before this point will not be written to the log
    logger.info('******************* TRACEABILITY UTILITY *******************')
    logger.info('Generating output to:\n\t%s' % (args.outputDir))
    
    errCode = validateArgs(args)
    if (0 != errCode):
        print ('Failed to parse command line arguments. View log for additional details.')
        exit(errCode)

    # export DOORS modules to CSV files
    if (True == args.EXPORT):
        errCode = exportDoorsModules(
            args.modules, 
            args.doorsUsr, 
            args.doorsPwd, 
            args.doorsServer,
            args.doorsView, 
            args.doorsExe, 
            args.outputDir)
        
        if (0 != errCode):
            print ('Failed to export DOORS modules. View log for additional details.')
            exit(errCode)

    # build requirements map from CSV files
    errCode, reqMap = buildReqMap(args.modules, args.outputDir)
    if (0 != errCode):
        print ('Failed to parse requirements modules. View log for additional details.')
        exit(errCode)
    
    if (True == args.checkSrcLinks):
        # get source code links
        for srcDir in args.srcDirs:
            parseSourceReqLinks(srcDir, args.outputDir, reqMap)
        # get model links
        for rpyFile in args.rpyFiles:
            parseRhapsodyModelLinks(rpyFile, reqMap)
            
    if (True == args.checkTestLinks):
        # get test links
        for tideDir in args.tideDirs:
            parseTideTestLinks(tideDir, args.outputDir, reqMap)
       
    if (True == args.TRACE):
        # generate traceability matrix
        logger.info('Generating traceability matrix:\n\t%s' % (os.path.join(args.outputDir, args.outfile + '.xlsx')))
        TraceabilityGenerator.generateTraceabilityMatrix(reqMap, args)
        
    if (True == args.JENKINS):
        # generate XML summary table for Jenkins
        generateJenkinsSummary(reqMap, args)
        pass
    
    if (True == args.REPORT):
        # generate report of missing requirements
        generateReport(reqMap, args)
    