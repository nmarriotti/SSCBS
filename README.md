# System Software Collector and Baseline Scanner (SSCBS)
Remotely extracts software package information. Creates a Microsoft Excel spreadsheet containing the findings. The generated output can then be used a baseline to compares findings and highlight software discrepancies such as:

1. Missing software
2. Version mismatch
3. Architecture issues
4. New unapproved software

## How to Use

------

### Creating a configuration

Configuration files specify the the hostname and IP address of the remote host as well as the credentials used to remotely access the system. Multiple configuration files can be created.

1. Open config/example.txt in a text editor
2. Specify the hostname and IP address within the __[hosts]__ section. For windows hosts prepend  __windows___ to the hostname
3. Provide the credentials within the __[credentials]__ section.
4. Below is an example configuration file containing both Windows and Linux hosts.

```
[credentials]
# Account used to access hosts
linux_username=yourusername
linux_password=yourpassword
windows_username=yourusername
windows_password=yourpassword

[hosts]
# Add hosts by specifying hostname=ip
# Prepend windows_ to all Windows hosts
ansible=1.1.1.1
windows_laptop=2.2.2.2
```

### Fetching Installed Software

```
python SSCBS.py
```

### Creating a Baseline

To create a baseline, simply copy an Excel file from the _Excel Files_ directory and paste it into _baselines_ directory making sure to rename it.

```
COPY: example.xlsx FROM: Excel Files\
To: baselines\  AS: example.baseline.xlsx
```

### Comparing Installed Software with a Baseline

Simply rerun SSCBS.py and if a baseline is located that matches the configuration filename then the scanner will start automatically once updated packages are retrieved.

```
 System Software Collector + Baseline Scanner for Windows/Linux - Version 1.3


    Using config file: example.txt

    Fetching packages, please wait...
       ansible => OK
                Additional Packages:
                        rubygem-io-console.0.4.2-36.el7.x86_64
                        rubygem-logging.1.8.2-1.el7.noarch
                        rubygem-apipie-bindings.0.0.13-1.el7.noarch
                        rubygem-kafo_parsers.0.0.5-1.el7.noarch
                        puppetlabs-release.22.0-2.noarch
                        ruby-irb.2.0.0.648-36.el7.noarch
                Version Mismatch:
                        gpg-pubkey 4bd6ec30-4c37bb40 installed; requires: a15703c6-5de96087
                        kernel 3.10.0-1062.18.1.el7 installed; requires: 3.10.0-1062.el7
                        kernel 3.10.0-1062.12.1.el7 installed; requires: 3.10.0-1062.el7
                        gpg-pubkey 352c64e5-52ae6884 installed; requires: a15703c6-5de96087
                        kernel 3.10.0-1127.10.1.el7 installed; requires: 3.10.0-1062.el7
                        gpg-pubkey f4a80eb5-53a7ff4b installed; requires: a15703c6-5de96087


    Save As (default 'output.xlsx'):

    File Location: C:\Users\nmarr\Documents\Development\RPM2Excel\SSCBS\Excel Files\output.xlsx
```