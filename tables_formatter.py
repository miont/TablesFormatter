# -*- coding: utf-8 -*-
"""
Форматирование таблиц (модификация размеров листа и прочее)

Для docx
"""
import fnmatch
import os
import os.path
import re
import argparse
import time
import logging
import papersize
import zipfile
import fnmatch
from decimal import Decimal
import math
import xml.etree.ElementTree
#import lxml
import linecache
import sys

## Constants
LOG_FILE_NAME = 'log/tables_formatter.log'    # Path to log-file
LOGGER_NAME = 'logger'                   # General logger name
DEFAULT_DEBUG_MODE = True
ORIENT_PORTRAIT = 'portrait'
ORIENT_LANDSCAPE = 'landscape'
WORD_DOCUMENT_XML_PATH = 'word/document.xml'

## Globals parameters
log = None           # Logger object

def input_args():
#    -or, --orient - ориентация листа (portrait/landscape)
#    -fmt, --format - формат листа (unsigned)
#    -f, --file - имя редактируемого файла
    """
        Input command line arguments
    """
    try:
        parser = argparse.ArgumentParser(description='Форматирование сводных таблиц АСТРА-НОВА')
        parser.add_argument('-or', '--orient', type=str, help='ориентация листа (portrait/landscape)')
        parser.add_argument('-fmt', '--format', type=str, help='формат листа (unsigned)')
        parser.add_argument('-f', '--file', type=str, help='имя редактируемого файла')
        parser.add_argument('--debug', dest='debug', action='store_true', required=False, help='режим отладки')
        parser.set_defaults(debug=DEFAULT_DEBUG_MODE)    
        return parser.parse_args()
    except Exception as e:
        raise ArgParserException(e)

def configure_logging(debug=False):
    """
    Logging initialization
    
    Args:
        debug (boolean): debug mode
    Returns:
        logger
    
    """
    try:
        logger = logging.getLogger(LOGGER_NAME)
    
        # Log to console        
        console_log_handler = logging.StreamHandler()
        logger.addHandler(console_log_handler)
        
        # Log to file
        if not os.path.exists('log'):
            os.makedirs('log')
        file_log_handler = logging.FileHandler(LOG_FILE_NAME, mode='w')
        logger.addHandler(file_log_handler)
        
        # Setup format
        formatter = logging.Formatter('%(message)s')
        file_log_handler.setFormatter(formatter)
        console_log_handler.setFormatter(formatter)
        
        # Level
        logger.setLevel('DEBUG' if debug else 'INFO')
        
        logger.propagate = False
        
        return logger
    except Exception as e:
        raise LoggerInitException(e)

def format_tables(file_path, sheet_format, sheet_orient=ORIENT_PORTRAIT):
    """
    Main procedure for tables formatting.
    
    Args:
         file_path: path to input file with tables to format
         sheet_orient: target sheet orientation
         sheet_format: target sheet format         
        
    Returns:
        none   
    """
    try:
        log.info('Extracting document ' + file_path)
        extract_path = file_path[:file_path.rfind('.')]
        extract_zip(file_path, extract_path)
        # Remove file
        #    os.remove(file_path)    
        
        log.info('Modifying format')
        # New size
        size = convert_to_word_units(get_paper_size(sheet_format, sheet_orient))

        # Modify word/document.xml
        modify_paper_format_document(os.path.join(extract_path, WORD_DOCUMENT_XML_PATH), size, sheet_orient)
        
        log.info('Archiving back')
        make_zipfile(extract_path, file_path)
        
    except Exception as e:
        raise TablesFormatterException(e)
        
def extract_zip(file, dest_dir):
    """
    Extracting zip archive
    
    Args:
        file: path to archive file
        dest_dir: path to directory where extract
    """
    zfile = zipfile.ZipFile(file, 'r')
    zfile.extractall(dest_dir)
    zfile.close()

def make_zipfile(source_dir, output_filename, include_root_dir=False):
    """
    Archives content of source directory in zip format
    
    Args:
        source_dir (str): full or relative path to the directory containing content needed to be archived.
        output_filename (str): name of generating archive file.
        include_root_dir (str): should root directory be created in archive.
    
    Returns:
        none
    """
    import zipfile, os.path
    method = zipfile.ZIP_DEFLATED  # method of compression
    source_dir = os.path.abspath(source_dir)    
    with zipfile.ZipFile(output_filename, 'w', method) as zf:
        for root, dirs, files in os.walk(source_dir):
            # add directory (needed for empty dirs)
            if root != source_dir or include_root_dir:
                zf.write(root, arcname=os.path.relpath(root,source_dir))
            # add files
            for file in files:
                filename = os.path.join(root,file)
                if os.path.isfile(filename): # regular files only
                    arcname = os.path.join(os.path.relpath(root,source_dir),file)
                    zf.write(filename, arcname)

def get_paper_size(sheet_format, sheet_orient, units='mm'):
    """
    Returns paper sizes for given format
    
    Args:
        sheet_format (str): format of paper
        sheet_orient (str): orientation
        units (str): units for size

    Returns:
        size (float tuple): sizes of paper
    
    """
    # Get size for given format
    size = papersize.parse_papersize(sheet_format, units)
    # Rotate if needed for given orientation
    size = papersize.rotate(size, papersize.PORTRAIT if sheet_orient.lower() == 'portrait' else papersize.LANDSCAPE)
    return map(float, size)
    
    
def convert_to_word_units(size, units='mm'):
    """
    """
    # Convert size to millimeters
    size = tuple(map(lambda d : papersize.convert_length(d, units, 'mm'), size))
    
    # Convert millimeters to twips
    mm_to_in = Decimal(1.0/25.4)
    in_to_twip = Decimal(1440)
    size = tuple(map(lambda d : math.ceil(d*mm_to_in*in_to_twip), size))
    return size
    
def modify_paper_format_document(file, size, orient):
    """
    """
    # Read namespaces
    ns = {}
    with open(file, 'r', encoding="utf8") as f:
        content = ''.join([line.strip() for line in f.readlines()])
        ns = read_namespaces(content)
    
    # Parse XML
    tree = xml.etree.ElementTree.parse(file)
    root = tree.getroot()
    
    # Look for tags w:pgSz and modify attributes 
    for elem in root.findall('.//w:pgSz', ns):
        elem.set('{%s}w'%(ns['w']), str(size[0]))
        elem.set('{%s}h'%(ns['w']), str(size[1]))
        elem.set('{%s}orient'%(ns['w']), orient)
        
    # Save
    tree.write(file)
        
def read_namespaces(xml_content):
    """
    Extracts namespaces from xml document as map
    
    Args:
        xml_content (str)
    Returns:
        ns (map): map containing (namespace - URL) pairs
    """
    
    ns = {}
    for m in re.finditer(r'xmlns:\s*(\w+)\s*="([^"]+)"', xml_content, flags=re.DOTALL):  # Not greedy match '+?'
        ns[m.group(1)] = m.group(2)
    return ns

def run_from_command_line():
    """
    Running application from command line        
    """
    global log
    try:
        args = input_args()
        log = configure_logging(args.debug)
        format_tables(args.file, args.format, args.orient)
    except ArgParserException as e:
        print('[ОШИБКА] Ошибка при парсинге аргументов командной строки')
        if DEFAULT_DEBUG_MODE:
            print(format_exception())
    except LoggerInitException as e:
        print('[ОШИБКА] Ошибка при настройке логирования: ' + str(e))
        if args.debug:
            log.error(format_exception())
    except TablesFormatterException as e:
        log.info('[ОШИБКА] Ошибка в процессе форматирования')
        log.debug(format_exception())
    except Exception as e:
        if log != None:
            log.info('[ОШИБКА]')
            log.debug(format_exception())
        else:
            print(format_exception())
    finally:
        if log != None:
            handlers = log.handlers[:]
            for handler in handlers:
                handler.close()
                log.removeHandler(handler)
        
# Exceptions
class ArgParserException(Exception):
    pass
class LoggerInitException(Exception):
    pass
class TablesFormatterException(Exception):
    pass

def format_exception():
    """
    https://stackoverflow.com/questions/14519177/python-exception-handling-line-number
    """
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    return ('EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))

        
if __name__ == '__main__':
    run_from_command_line()

