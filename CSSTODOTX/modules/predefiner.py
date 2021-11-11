"""Predefined functions for my code such as:
printError
printInfo
printDebug
printWarning
"""

__version__ = '0.1'
__author__ = 'Samuel Voss'
__copyright__ = 'Copyright 2021 Samuel Voss'

__license__ = 'MIT'
__maintainer__ = 'Samuel Voss'
__email__ = 'samvoss69@gmail.com'
__status__ = 'Development'

from time import process_time as ptime
from colorama import Fore, Style, init
init()

def printError(message):
    """Prints an error message to the console with ANSI colors."""
    try:
        message = str(message)
    except Exception as e:
        print(f"{Fore.RED}{str(ptime())}: [ERROR]{Style.RESET_ALL} {Fore.WHITE}" + str(e) + Style.RESET_ALL)

    print(f"{Fore.RED}{str(ptime())}: [ERROR]{Style.RESET_ALL} {Fore.WHITE}" + message + Style.RESET_ALL)

def printInfo(message):
    """Prints an info message to the console with ANSI colors."""
    try:
        message = str(message)
    except Exception as e:
        print(f"{Fore.RED}{str(ptime())}: [ERROR]{Style.RESET_ALL} {Fore.WHITE}" + str(e) + Style.RESET_ALL)
    print(f"{Fore.CYAN}{str(ptime())}: [INFO]{Style.RESET_ALL} {Fore.WHITE}" + message + Style.RESET_ALL)

def printDebug(message):
    """Prints a debug message to the console with ANSI colors."""
    try:
        message = str(message)
    except Exception as e:
        print(f"{Fore.RED}{str(ptime())}: [ERROR]{Style.RESET_ALL} {Fore.WHITE}" + str(e) + Style.RESET_ALL)
    print(f"{Fore.YELLOW}{str(ptime())}: [DEBUG]{Style.RESET_ALL} {Fore.WHITE}" + message + Style.RESET_ALL)

def printWarning(message):
    """Prints a warning message to the console with ANSI colors."""
    try:
        message = str(message)
    except Exception as e:
        print(f"{Fore.RED}{str(ptime())}: [ERROR]{Style.RESET_ALL} {Fore.WHITE}" + str(e) + Style.RESET_ALL)
    print(f"{Fore.LIGHTRED_EX}{str(ptime())}: [WARNING]{Style.RESET_ALL} {Fore.WHITE}" + message + Style.RESET_ALL)

def printSuccess(message):
    """Prints a success message to the console with ANSI colors."""
    try:
        message = str(message)
    except Exception as e:
        print(f"{Fore.RED}{str(ptime())}: [ERROR]{Style.RESET_ALL} {Fore.WHITE}" + str(e) + Style.RESET_ALL)
    print(f"{Fore.GREEN}{str(ptime())}: [SUCCESS]{Style.RESET_ALL} {Fore.WHITE}" + message + Style.RESET_ALL)