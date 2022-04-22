# geocoder
Geocoder program (Match Admin Boundaries) used to match address or location names in a spreadsheet file to administrative boundaries in a shapefile. This application supports Microsoft Excel or CSV/Comma Separated Value spreadsheets and it can do fuzzy-matching. This tool was developed while volunteering with the Canadian Red Cross.

The geocoder application requires that you have Python Version 3.7 or newer. It also uses Python’s PIP package manager, which is included in Python. 

### Installation Instructions for Windows Users:

* pip install wheel
* pip install pipwin
* pipwin install numpy
* pipwin install pandas
* pipwin install shapely
* pipwin install gdal
* pipwin install fiona
* pipwin install pyproj
* pipwin install six
* pipwin install rtree
* pipwin install geopandas
* pip install openpyxl
* pip install unidecode
* pip install wxpython
* pip install pathvalidate
* pip install bs4
* pip install thefuzz
* Optional Installation to Install the Fast Version of thefuzz: pip install thefuzz[speedup]  See https://pypi.org/project/thefuzz/ or Google Search: thefuzz[speedup] for more info

### Installation Instructions for Mac and Linux Users:

* pip install gdal
* pip install fiona 
* pip install geopandas
* pip install openpyxl
* pip install unidecode
* pip install wxpython
* pip install pathvalidate
* pip install bs4
* pip install thefuzz or pip install thefuzz[speedup] for the faster version of thefuzz.
* See https://pypi.org/project/thefuzz/ or Google Search: thefuzz[speedup] for more info

### Installation Instructions for Anaconda Users:
* Create [conda environment first.](https://stackoverflow.com/questions/61415344/cant-install-geopandas-with-anaconda-because-of-conflicts)
* Install python libraries as detailed in "Installation Instructions for Mac and Linux Users" section above.

### Intructions for Operating the Geocoder Graphical User Interface (GUI) Application:

1. Make sure you have the two Python files match_admin_boundaries_core.py and match_admin_boundaries_gui.py in the same folder,
e.g. c:\temp folder.
2. Start the GUI in your Python environment: python match_admin_boundaries_gui.py
3. You’ll see the GUI application.
4. Follow these steps to use the GUI app:
   1. Click the Maximize icon on top right hand corner, to enlarge the window first.
   2. Click "Choose Spreadsheet File" Button and select spreadsheet file (.csv, .xls, and .xlsx files only)
   3. Click the "Choose Administratve Boundary Shapefile" button and select the administrative boundary shapefile.
   4. Click the "Run the Matching Process" button and follow the GUI application prompts.

### Intructions for Operating the Geocoder Console Application:

The Console version of the Geocoder app is a "bare-bones" application that might work better than the GUI application on some computers.

1. Make sure you only have the Python file match_admin_boundaries_core.py in a folder, e.g. c:\temp folder.

2. Start your Python environment (Windows Python prompt, or Python prompt in Linux or Mac or Anaconda prompt). [Info for setting up Windows command prompt to run Python.](https://www.geeksforgeeks.org/how-to-set-up-command-prompt-for-python-in-windows10)

3. The console program requires 3 parameters, spreadsheet_file, admin_boundaries_file, and match_type. Run console commands as shown below and follow the console application prompts.

__4. Sample console commands:__

`python match_admin_boundaries_core.py --spreadsheet_file "c:\temp\AddressData.xlsx" --admin_boundaries_file "c:\temp\hnd_admbnda_adm3_sinit_20161005.shp" --match_type regular`

`python match_admin_boundaries_core.py --spreadsheet_file "c:\temp\AddressData.csv" --admin_boundaries_file "c:\temp\hnd_admbnda_adm3_sinit_20161005.shp" --match_type fuzzy` 

__5. Sample ABBREVIATED console commands:__

`python match_admin_boundaries_core.py -s "c:\temp\AddressData.xlsx" -a "c:\temp\hnd_admbnda_adm3_sinit_20161005.shp" -m regular`

`python match_admin_boundaries_core.py -s "c:\temp\AddressData.xlsx" -a "c:\temp\hnd_admbnda_adm3_sinit_20161005.shp" -m fuzzy`
