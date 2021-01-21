#=============================================================================
# Title: System Software Collector + Baseline Scanner (SSCBS)
# Version: 1.3
# Description: Extracts installed packages from Windows/Linux hosts and
#              compares each package with approved baselines
# Author: Nick Marriotti
# Date: 2/11/2020
#=============================================================================
import paramiko, time
import openpyxl, sys, os
from subprocess import check_output


configurations = {}
idx = ""
saveLocation = os.path.join(os.getcwd(), "Excel Files")
configfile = ""


# Only prints the banner
def printTitle():
    os.system("cls")
    title = '''\n\033[1;36;40m System Software Collector + Baseline Scanner for Windows/Linux - Version 1.3\033[1;37;40m'''
    print(title)
    print("")
    print("")


class Scanner:
    def __init__(self, configfile):
        self.configfile = configfile
        self.baseline_directory = os.path.join(os.getcwd(), "baselines")
        self.loaded = False
        self.rolodex = {}
        self.baseline_file = ""
        self.hostname = ""
        self.load()

        # Used for metric purposes
        self.additional = []
        self.missing = []
        self.version_mismatch = []
        self.arch_mismatch = []
    
    # If a baseline file exists for the current system, load it    
    def load(self):
        # Build expected baseline filename
        self.baseline_file = os.path.join(self.baseline_directory, configfile.replace(".txt",".baseline.xlsx").strip())

        # Does baseline exist in baseline directory?
        if os.path.basename(self.baseline_file) in os.listdir(self.baseline_directory):
            #print("Baseline loaded")
            self.loaded = True

    # Builds a dictionary from baseline packages used to compare with.
    # packages are distributed into alphabetic keys to speed up search
    # process and are removed if the current package satisfies baseline
    # requirements.
    def build(self):
        workbook = openpyxl.load_workbook(filename=self.baseline_file)
        sheetnames = workbook.sheetnames
        for sheet in sheetnames:
            # Add each sheet as rolodex key
            if not sheet in self.rolodex.keys():
                self.rolodex[sheet] = {}
                current_sheet = workbook[sheet]
                # Get each row from baseline
                for row in current_sheet.iter_rows(values_only=True):
                    try:
                        # Package information
                        package_name = row[0].strip()
                        package_version = row[1].strip()
                        package_arch = ""

                        # Architecture only available in Linux 
                        if not "windows" in self.hostname:
                            package_arch = row[2].strip()

                        # First letter of package name for rolodex organization
                        rolodex_index = package_name[:1].upper()
                        
                        # DEBUG - Show alphanumeric key package will be stored in
                        #print("rolodex index: {}".format(rolodex_index))

                        if not rolodex_index in self.rolodex[sheet]:
                            # Add new alphabetic key 
                            self.rolodex[sheet][rolodex_index] = {"packages":{}}
                        # Add each package to its correct letter index
                        self.rolodex[sheet][rolodex_index]["packages"][package_name] = {"version":package_version, "arch":package_arch}
                    except Exception as e:
                        #print(e)
                        pass

        # DEBUG - Print complete rolodex scanner will use for comparisons
        #print(self.rolodex)

    def isLoaded(self):
        return self.loaded

    def setHostnameAndBuild(self, hostname):
        self.hostname = hostname
        self.build()

    # Compares packages installed on the host with the baseline
    def start(self, packages):
        for package in packages:
            try:
                package_name = ""
                package_version = ""
                package_arch = ""

                if "windows" in self.hostname:
                    package_name, package_version = package.strip().split(" ")
                else:
                    package_name, package_version, package_arch = package.strip().split(" ")

                idx = package_name.strip()[:1].upper()
                if idx in self.rolodex[self.hostname].keys():
                    # Package check
                    if package_name in self.rolodex[self.hostname][idx]["packages"].keys():
                        # Version check
                        if package_version == self.rolodex[self.hostname][idx]["packages"][package_name]["version"]:
                            # Architecture check
                            if package_arch == self.rolodex[self.hostname][idx]["packages"][package_name]["arch"]:
                                # Valid package found, remove from rolodex
                                del self.rolodex[self.hostname][idx]["packages"][package_name]
                            else:
                                # Architecture invalid
                                required_arch = self.rolodex[self.hostname][idx]["packages"][package_name]["arch"]
                                self.arch_mismatch.append("{} {} installed; requires: {}".format(package_name, package_version, required_version))
                        else:
                            # Version does not match
                            required_version = self.rolodex[self.hostname][idx]["packages"][package_name]["version"]
                            self.version_mismatch.append("{} {} installed; requires: {}".format(package_name, package_version, required_version))
                    else:
                        # New package that is not in the baseline
                        self.additional.append("{}.{}.{}".format(package_name, package_version, package_arch))
                else:
                    # New package that is not in the baseline
                    # Rolodex did not include this index
                    self.additional.append("{}.{}.{}".format(package_name, package_version, package_arch))                   
            except Exception as e:
                #print(package)
                #print(e)
                pass

        # DEBUG- Items should be empty if everything matched up with the baseline
        #print("\nMissing Packages")
        #print(self.missing)
        #print("\nAdditional Packages")
        #print(self.additional)
        #print("\nVersion Mismatch")
        #print(self.version_mismatch)
        #print("\nArch Mismatch")
        #print(self.arch_mismatch)
        #print("")
        #print(self.rolodex)
        #print("")

        # Display findings to the user
        self.results()

        # Reset
        self.__init__(self.configfile)



    # Count the number of packages remaining in the rolodex
    def calculate_remaining_packages(self):
        i = 0
        for ascii in range(65, 90):
            try:
                idx = str(chr(ascii))
                for key, value in self.rolodex[hostname][idx]["packages"].items():
                    if key:
                        i += 1
                        #print("Missing: {}".format(key))
                        self.missing.append("[{}, {}. {}]".format(key, value["version"], value["arch"]))
            except Exception as e:
                return i
        return i



    # Display results to the user
    def results(self):
        num_additional = len(self.additional)
        num_version = len(self.version_mismatch)
        num_arch = len(self.arch_mismatch)
        num_missing = int(self.calculate_remaining_packages())

        if num_missing > 0:
            print("\t\t\033[1;31;40mMissing Packages:")
            for package in self.missing:
                print("\t\t\t{}".format(package))

        if num_additional > 0:
            print("\t\t\033[1;31;40mAdditional Packages:")
            for package in self.additional:
                print("\t\t\t{}".format(package))

        if num_version > 0:
            print("\t\t\033[1;31;40mVersion Mismatch:")
            for package in self.version_mismatch:
                print("\t\t\t{}".format(package))

        if num_arch > 0:
            print("\t\t\033[1;31;40mArch Mismatch:")
            for package in self.arch_mismatch:
                print("\t\t\t{}".format(package))



# Populate menu of all files in config directory
def display_available_systems():
    global configurations
    configurations = {}
    path = os.path.join(os.getcwd(), "config")
    i = 1
    for file in os.listdir(path):
        if not file == "README":
            configurations[str(i)] = str(file)
            i += 1
    if len(configurations) > 0:
        for key, value in configurations.items():
            print("       {}. {}".format(key, value))
        return True
    else:
        return False


        
# Ask user to select a config file            
def setTarget():
    global idx
    target = input("\n    Selection (1-{}) => ".format(len(configurations)))
    idx = str(target)
    return idx



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
    if not setTarget() in configurations.keys():
        # User selected an invalid option
        return "Invalid"

    # Got a valid config file
    return configurations[idx]

 

# Process the config file
def load_answer_file(f):
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
        stdin, stdout, stderr = client.exec_command('rpm --queryformat "%{NAME} %{VERSION}-%{RELEASE} %{ARCH}\n" -qa | sort -n\n')
        packages = stdout.read().decode(encoding='utf-8').split("\n")

        '''for idx in range(0, len(packages)-1):
            parts = packages[idx].split(" ")
            package = "{0}-{1}.{2}.rpm".format(parts[0], parts[1], parts[2])
            file_download = "{0}-{1}".format(parts[0], parts[1])
            stdin, stdout, stderr = client.exec_command("echo '{0}' | sudo -S yum reinstall --downloadonly --downloadonly --downloaddir=. {1} && md5sum {2} && rm -f {2}".format(credentials['linux_password'], file_download, package), get_pty=True)
            output = stdout.read().decode(encoding='utf-8').split("\n")
            md5 = output[-2]
            packages[idx] += " {0}".format(md5)'''
        
        for idx in range(0, len(packages)-1):
            parts = packages[idx].split(" ")
            package = "{0}-{1}.{2}.rpm".format(parts[0], parts[1], parts[2])
            packages[idx] += " {0}".format(package)

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


# Returns list of necessary columns (A1, A2, A3, etc...)
def getColumns(length, row):
    columns = []
    for ascii in range(97, 97 + length):
        letter = chr(ascii).upper()
        col = "{}{}".format(letter, row)
        columns.append(col)
    return columns



# Start here
while True:
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

            package = package.split(" ")
            length = len(package)

            # Get columns
            columns = getColumns(length, row)

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