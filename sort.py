import os, sys, json
from pathlib import Path

for index, arg in enumerate(sys.argv):
    if arg in ["-p", "--path"]:
        try:
            os.chdir(sys.argv[index + 1])
        except IndexError:
            raise SystemExit("Folder path not provided or misplaced, it should be the argument after the '--path' or '-p' flag\nExiting prematurely")
        except: # Ideally this would not be a bare except, BUT there's so many that could reasonably appear
                # that I cannot find the willpower to handle manually
            raise SystemExit(f"Folder path '{sys.argv[index + 1]}' is incorrect, please validate or omit\nExiting prematurely")

CONFIGFILE = os.path.join(".", "config.json")
for index, arg in enumerate(sys.argv):
    if arg in ["-u", "--usejson"]:
        try:
            CONFIGFILE = sys.argv[index + 1]
        except IndexError:
            raise SystemExit("File path not provided or misplaced, it should be the argument after the '--usejson' or '-u' flag\nExiting prematurely")
        except: # Ideally this would not be a bare except, BUT there's so many that could reasonably appear
                # that I cannot find the willpower to handle manually
            raise SystemExit(f"File path '{sys.argv[index + 1]}' is incorrect, please validate or omit\nExiting prematurely")

# Some configuration values and/or false positives
nojson = True
directories = None
exclude = [os.path.basename(__file__), "config.json"]
want_help = "-h" in sys.argv or "--help" in sys.argv
make_json = "-m" in sys.argv or "--makejson" in sys.argv
verbose = "-v" in sys.argv or "--verbose" in sys.argv
overwrite = "-o" in sys.argv or "--overwrite" in sys.argv

# Yes, all of this is to either get data from the config file OR use the preprogrammed defaults
# ALL OF THIS
if os.path.exists(CONFIGFILE):
    try:
        with open(CONFIGFILE, "r") as f:
            data = json.load(f)
            if isinstance(data, dict):
                directories = data.get("directories", None)
                raw_exclude = data.get("exclude", None)
                verbose = data.get("verbose", verbose)
                overwrite = data.get("overwrite", overwrite)
                folderpath = data.get("folder_path", None)
                if isinstance(raw_exclude, list):
                    exclude = raw_exclude if raw_exclude else exclude
                if isinstance(directories, dict) and directories is not None:
                    nojson = False
                if folderpath:
                    os.chdir(folderpath)
    except json.JSONDecodeError:
        print("Configuration file is malformed, please validate\nDefaulting to preset configuration")
        nojson = True
nojson = True if "--nojson" in sys.argv or "-n" in sys.argv else nojson

if nojson:
    directories = {
        "Excel": ["xls", "xlsx", "xlsm", "xlsb", "xltx", "xltm", "xla", "xlam", "xlk"],
        "PDF": ["pdf"],
        "Outlook": ["msg", "pst", "ost", "oft"],
        "Word": ["doc", "docx", "docm", "dot", "dotx", "dotm", "asd", "wbk"],
        "PowerPoint": ["ppt", "pptx", "pps", "ppsx", "ppsm", "pptm", "pot", "potx", "potm", "ppa", "ppam"],
        "OneNote": ["one", "onetoc2"],
        "Visio": ["vsd", "vsdx", "vsdm", "vss", "vssx", "vst", "vstx"],
        "Project": ["mpp", "mpt"],
        "Access": ["mdb", "accdb", "accde", "accdt", "accda", "laccdb", "ldb"]
    }

# Strap in for a huge text wall
if want_help:
    print("""
How to use:


There are two main ways to use this sorter, both good in different situations.

Either drop this file in the folder in which you want to sort and run "python3 sort.py"
OR
Run this: "python3 sort.py --path <folder_to_sort>" where <folder_to_sort> is the path to said folder
(You can replace "python3" with "python" if you're certain it works, sometimes it doesn't)


To configure:


If you want different functionality in the sorter, such as a different folder setup or to know what's moved where, run this to start:
"python3 sort.py --makejson"

This forces the sorter to not sort for one run and instead give you the default configuration in a file to edit. To edit, just follow the format set up.
You may add or remove the following sections as you wish:

    directories
    exclude
    verbose
    overwrite
    folder_path

The directories section is how the sorter sorts files. The exclude section is where you add names of files (including extension) that you don't want moved.
The verbose section is whether you want the sorter to tell you what and where every file is moved, it is functionally the same as the --verbose option.
Note: if a section is missing, it is replaced with the default version of that section, verbose and overwrite are false (false or 0) by default.
Note: if a configuration file specifies a command-line argument as true (true or 1), it overrides the command-line interface.


If you want a more advanced list of commands, here it is:

    -p, --path     | Used to specify the path to a folder in need of sorting.
    -v, --verbose  | Used to see what files are moved and where; also configurable in a configuration file.
    -o, --overwrite| Used to confirm any file overwrite attempt by the sorter; also configurable in a configuration file; CAUTION.
    -h, --help     | Used to display this message.
    -n, --nojson   | Used to tell the sorter to use the preset configuration regardless of whether a configuration file is present.
    -u, --usejson  | Used to point the sorter to a non-standard configuration file; useful when different folders have different requirements.
    -m, --makejson | Used to tell the sorter to not sort a folder this time and instead place a preset configuration into a new JSON file.""")
# The following format is used to configure the sorter, but it is optional for basic uses
elif make_json:
    config_data = {
        "directories": {
            "Excel": ["xls", "xlsx", "xlsm", "xlsb", "xltx", "xltm", "xla", "xlam", "xlk"],
            "PDF": ["pdf"],
            "Outlook": ["msg", "pst", "ost", "oft"],
            "Word": ["doc", "docx", "docm", "dot", "dotx", "dotm", "asd", "wbk"],
            "PowerPoint": ["ppt", "pptx", "pps", "ppsx", "ppsm", "pptm", "pot", "potx", "potm", "ppa", "ppam"],
            "OneNote": ["one", "onetoc2"],
            "Visio": ["vsd", "vsdx", "vsdm", "vss", "vssx", "vst", "vstx"],
            "Project": ["mpp", "mpt"],
            "Access": ["mdb", "accdb", "accde", "accdt", "accda", "laccdb", "ldb"]
        },
        "exclude": [
            os.path.basename(__file__),
            "config.json"
        ],
        "verbose": False,
        "overwrite": False,
        "folder_path": None
}
    try:
        with open(CONFIGFILE, "w") as f:
            json.dump(config_data, f, indent=4)
    except:
        print(f"Unable to create file, writing contents to command line\nFormat and place into file named 'config.json'\n\n{config_data}")
else:
    for file in os.listdir():
        if os.path.isdir(file) or file in exclude:
            continue
        ext = Path(file).suffix.lstrip(".").lower()
        moved = False
        for directory, extension in directories.items():
            if ext in extension:
                try:
                    os.makedirs(directory, exist_ok=True)
                    newpath = os.path.join(directory, file)
                    # This hack is for the overwrite option and is far from optimal             Edit: it got worse, don't let me read the os library docs ever again
                    if os.path.exists(newpath):
                        if overwrite:
                            os.replace(file, newpath)
                        elif input(f"Duplicate file named '{file}' found at '{newpath}'\nOverwrite file (Y/N)? ").lower() in ["y", "yes"]:
                            os.replace(file, newpath)
                    else:
                        os.replace(file, newpath)
                    moved = True
                    # This is literally just for when the folder gets huge
                    if verbose: print(f"File '{file}' moved to '{newpath}'")
                    break
                except PermissionError:
                    print(f"File '{file}' cannot be moved to '{newpath}', run as administrator or check folder permissions")
                except OSError as e:
                    print(e)
        if not moved:
            try:
                os.makedirs("Misc", exist_ok=True)
                newpath = os.path.join("Misc", file)
                if os.path.exists(newpath):
                    if overwrite:
                        os.replace(file, newpath)
                    elif input(f"Duplicate file named '{file}' found at '{newpath}'\nOverwrite file (Y/N)? ").lower() in ["y", "yes"]:
                        os.replace(file, newpath)
                else:
                    os.replace(file, newpath)
                if verbose: print(f"File '{file}' moved to '{newpath}'")
            except PermissionError:
                print(f"File '{file}' cannot be moved to '{newpath}', run as administrator or check folder permissions")
            except OSError as e:
                print(e)
