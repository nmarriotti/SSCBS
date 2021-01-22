#=============================================================================
# Title: System Software Collector + Baseline Scanner (SSCBS)
# Version: 1.3.1
# Description: Extracts installed packages from Windows/Linux hosts and
#              compares each package with an approved baseline
# Author: Nick Marriotti
# Date: 1/21/2020
#=============================================================================
import paramiko, time
import openpyxl, sys, os
from subprocess import check_output
from modules.scanner import Scanner

# Global variables
configurations = {}
idx = ""
saveLocation = os.path.join(os.getcwd(), "Excel Files")
# Stores config file to use
configfile = ""


# Only prints the banner
def printTitle():
    os.system("cls")
    title = '''\n\033[1;36;40m System Software Collector + Baseline Scanner for Windows/Linux - Version 1.3\033[1;37;40m'''
    print(title)
    print("")
    print("")


def display_available_systems():
    ''' Prints a numbered menu of all files in the config directory '''
    global configurations
    configurations = {}
    # Build an absolute path to config directory
    path = os.path.join(os.getcwd(), "config")
    # Counter
    i = 1
    # Assign files to the configurations dict
    for file in os.listdir(path):
        if not file == "README":
            configurations[str(i)] = str(file)
            i += 1
    if len(configurations) > 0:
        # Print config files 
        for key, value in configurations.items():
            print("       {}. {}".format(key, value))
        return True
    else:
        return False

        
# Ask user to select a config file            
def chooseConfigFile():
    target = input("\n    Selection (1-{}) => ".format(len(configurations)))
    return str(target)


# Menu to prompt for config file
def display_menu():
    printTitle()
    print("    Select a configuraton file:\n")

    # Make sure config files are available
    while not display_available_systems():
        print("       \033[1;31;40m*** No systems present! ***\033[1;37;40m")
        input("\n    Press any key to refresh.")
        return "Invalid"

    # Compare selection with available options
    global idx
    global configurations

    # Config file user wants to load
    idx = chooseConfigFile()

    if not idx in configurations.keys():
        # User selected an invalid option
        return "Invalid"

    # Got a valid config file
    return configurations[idx]

 
# Process the config file
def load_answer_file(f):
    ''' 
    DESCRIPTION:
        Extact host information from the select config file
    PARAMS:
        f = config file
    RETURN:
        credentials
        Host IPs
    '''
    hosts = {}
    credentials = {}
    _hosts = False
    _credentials = False
    try:
        with open(f, 'r') as infile:
            data = infile.readlines()
            for line in data:
                line = line.strip()
                # ignore comments
                if line[:1] == "#":
                    continue
                if "[credentials]" in line:
                    _credentials = True
                    _hosts = False
                    pass
                elif "[hosts]" in line:
                    _hosts = True
                    _credentials = False
                    pass

                try:
                    line = line.strip().split("=")
                    if _credentials:
                        credentials[line[0]] = line[1]
                    elif _hosts:
                        if not line[0] in hosts.keys():
                            hosts[line[0]] = line[1]
                except Exception as e:
                    pass

        # Write Windows answer file
        with open('windows//answerfile.txt', 'w') as answerfile:
            answerfile.write('username={}\n'.format(credentials["windows_username"]))
            answerfile.write('password={}'.format(credentials["windows_password"]))

        return credentials, hosts
    except Exception as e:
        print(e)
        exit(1)


# Append the Windows IP to the Powershell answerfile
def addIpToAnswerFile(ip):
    ''' 
    DESCRIPTION:
        Writes Windows host information to the windows answer file
    PARAMS:
        ip = IP address of Windows host
    '''
    replaceIp = False
    answerfile = open('windows//answerfile.txt', 'r')
    data = answerfile.readlines()

    # Collect information that is currently in the answerfile
    for idx in range(0, len(data)):
        data[idx] = data[idx].strip()
        if "ComputerName" in data[idx]:
            replaceIp = True

    answerfile.close()

    # Remove the previous IP since it will be replaced
    if replaceIp:
        data = data[:len(data)-1]

    # Append new Windows IP to list
    data.append('ComputerName={}'.format(ip))

    # Rewrite the answerfile
    answerfile = open('windows//answerfile.txt', 'w')
    for line in data:
        answerfile.write('{}\n'.format(line))
    answerfile.close()


# Add file extension if missing, or use default filename
def fixFilename(filename):
    ''' 
    DESCRIPTION:
        Adds a .xlsx extension to the output filename if missing
    PARAMS:
        filename = Output filename
    RETURN:
        string : Absolute path to output file with .xlsx extension
    '''
    if filename == "":
        return os.path.join(saveLocation, "output.xlsx")
    if "." in filename:
        temp = filename.strip().split(".")
        extension = str(temp[len(temp)-1]).lower()
        if "xlsx" in extension or "xls" in extension:
            return os.path.join(saveLocation, filename)
        else:
            return os.path.join(saveLocation, "{}{}".format(filename, ".xlsx"))
    else:
        return os.path.join(saveLocation, "{}{}".format(filename, ".xlsx"))


# Connect to each host and collect RPM package listing    
def getPackages(credentials, hostname, ip):
    ''' 
    DESCRIPTION:
        Connects to host and extracts a list of installed packages
    PARAMS:
        credentials = username and password of host
        hostname = hostname of the host
        ip = IP address of host
    RETURN:
        list of installed packages
    '''
    sys.stdout.write("\033[1;32;40m       {} => ".format(hostname))
    sys.stdout.flush()

    # LINUX
    if not "windows" in hostname:
        # Connect to host via SSH
        try:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(ip, 22, credentials["linux_username"], credentials["linux_password"])
        except:
            sys.stdout.write("\033[1;31;40mConnection attempt failed\033[1;37;40m\n")
            return ["Failed"]

        # Grab output
        stdin, stdout, stderr = client.exec_command('rpm --queryformat "%{NAME} %{VERSION}-%{RELEASE} %{ARCH}\n" -qa\n')
        packages = stdout.read().decode(encoding='utf-8').split("\n")

        # Close the SSH connection
        client.close()

        sys.stdout.write("OK\n")

        # Return installed packages
        return packages[0:len(packages)-1]

    # WINDOWS
    # Requires Windows Remote Management (WinRM) to be enabled and properly configured
    else:
        # Execute Powershell script to obtain installed applications
        packages = check_output(["powershell.exe", "windows\\extract_packages.ps1"]).decode(encoding='utf-8').split("\n")
        if "ERROR" in packages:
            # Unable to connect to Windows host
            sys.stdout.write("\033[1;31;40mConnection attempt failed\033[1;37;40m\n")
            return ["Failed"]
        else:
            # Obtained Windows Applications
            sys.stdout.write("OK\n")
            return packages[0:len(packages)-1]


# Returns list of necessary columns (A1, B1, C1, etc...)
def getColumns(length, row):
    ''' 
    DESCRIPTION:
        Determines number of columns needed to write data to Excel workbook
    PARAMS:
        length = number of cells to write a a given row
        row = Number of row
    RETURN:
        list of Excel row/column identifiers
    '''
    columns = []
    for ascii in range(97, 97 + length):
        letter = chr(ascii).upper()
        col = "{}{}".format(letter, row)
        columns.append(col)
    return columns


# Start here
while True:
    # Prmopt user to select a configuration file
    configfile = display_menu()

    # Restart if invalid config file selected
    if configfile == "Invalid":
        continue

    # Get list of hosts
    credentials, hosts = load_answer_file(os.path.join(os.path.join(os.getcwd(), "config"), configfile))

    # Create excel object
    workbook = openpyxl.Workbook()
    # Remove default sheet
    workbook.remove(workbook.active)

    printTitle()
    print("    Using config file: {}\n".format(configfile))
    print("    Fetching packages, please wait...")

    # Scanner object to compare packages against baseline
    scanner = Scanner(configfile)

    # Iterate through all hosts
    for hostname, ip in hosts.items():
        
        # Add this IP to Windows answer file
        if "windows" in hostname:
            addIpToAnswerFile(ip)

        # Set current sheet and title
        sheet = workbook.create_sheet(hostname)
        sheet.title = hostname

        # Grab packages
        packages = getPackages(credentials, hostname, ip)

        # Starting column number for each row
        row = 1

        # Only scan if a baseline file is present
        if scanner.isLoaded() and not packages[0] == "Failed":
            scanner.setHostnameAndBuild(hostname)
            scanner.start(packages)
        
        # Format each package to be written to xlsx
        for package in packages:
            # Create list from package string
            package = package.split(" ")

            # Get columns
            columns = getColumns(len(package), row)

            # Go to next row
            row += 1

            # Map package data to corresponding column
            for idx in range(0, len(columns)+1):
                try:
                    # Write the data to workbook
                    sheet[columns[idx]] = package[idx]
                except Exception as e:
                    break

    print("\033[1;37;40m")

    # Save the file
    try:
        saveas = input("\n    Save As (default 'output.xlsx'): ")
        saveas = fixFilename(saveas)
        try:
            os.mkdir("Excel Files")
        except Exception as e:
            pass
        workbook.save(filename=saveas)
        print("\n    File Location: \033[1;32;40m{}\033[1;37;40m".format(saveas))
    except Exception as e:
        print(e)

    input("\n\nPress any key to return to main menu")

    # Rinse and repeat
    idx = ""
    configfile = ""