import geopandas
import pandas
import numpy
import platform
from os import path, mkdir
from unidecode import unidecode
import datetime
import re
from collections import OrderedDict, namedtuple
from argparse import ArgumentParser, RawDescriptionHelpFormatter
from pathvalidate import sanitize_filepath
from thefuzz import fuzz, process
from bs4 import UnicodeDammit

'''This module provides the core logic, i.e. the Model, for the GUI & console versions of the match_admin_boundaries 
geocoder application. '''


class PromptMessages(object):
    # This class stores prompt messages for reuse by console and GUI versions of the app
    # Class variable can be changed without creating instance
    _argument = None

    @property
    def argument(self):
        return type(self)._argument

    @argument.setter
    def argument(self, val):
        type(self)._argument = val

    ''' example call:
    p = PromptMessages()
    p.argument = 'click OK button'
    fuzzy_caption = p.fuzzy_caption'''

    @property
    def fuzzy_caption(self):
        return '\nEnter a number between 1 to 100 as a fuzzy match cut-off score and {0}.' \
               '\nIt\'s recommended that you enter a cut-off score between 50 and 100 for' \
               '\nmore accurate fuzzy match results.\n\n50 would be a flexible text match criterion,' \
               '\nand 100 would be the highest fuzzy match criterion.'.format(self.argument)

    @property
    def epsg_caption(self):
        return '\nA VALID EPSG CODE IN PROJECTED COORDINATES IS REQUIRED TO GENERATE THE SHAPEFILE' \
               '\nYou may refer to a website or GIS software for a valid PROJECTED coordinate system EPSG code' \
               '\nA valid EPSG code should contain 4 or 5 numerical digits, e.g. 3857 ' \
               '\nis Spherical Mercator projected coordinate system.' \
               '\nExample website to check for EPSG codes: https://epsg.io' \
               '\nPlease enter a valid EPSG Code in the text box and {0}.'.format(self.argument)


class DataUtility:

    @staticmethod
    def filter_row(dataframe, col_name, text_filter, spreadsheet_data):
        # returns a numpy ndarray
        # Only apply unidecode on Western European/Latin type languages
        if spreadsheet_data.encoding in spreadsheet_data.western_europe_encodings:
            dataframe[col_name] = DataUtility.remove_accented_char(dataframe[col_name])
            dataframe[col_name] = dataframe[col_name].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)
            filter = dataframe[col_name] == unidecode(text_filter).lower().strip()
        else:
            dataframe[col_name] = dataframe[col_name].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)
            filter = dataframe[col_name] == text_filter.lower().strip()

        # Yes Correct June 3 2021
        if len(dataframe[filter]) > 0:
            print('Filtered row with first column id of {0}'.format(dataframe[filter].values[0][0]))
            return dataframe[filter].values[0]

    @staticmethod
    def is_string_match(cell_text, column_series, spreadsheet_data):
        """Match text in spreadsheet cell against all values in Administrative level column
        column_series already has stripped white space """

        # Only remove accents on Western European/Latin type languages
        if spreadsheet_data.encoding in spreadsheet_data.western_europe_encodings:
            column_series = DataUtility.remove_accented_char(column_series)

        # Can not apply string comparison on GeometryDtype/geometry column
        if isinstance(column_series, geopandas.geoseries.GeoSeries) and isinstance(column_series.values,
                                                                                   geopandas.array.GeometryArray):
            return False
        elif pandas.api.types.is_string_dtype(column_series):
            if spreadsheet_data.encoding in spreadsheet_data.western_europe_encodings:
                return (unidecode(str(cell_text).lower().strip()) == column_series.str.lower()).any()
            else:
                return (str(cell_text).lower().strip() == column_series.str.lower()).any()
        else:
            return False

    @staticmethod
    def remove_accented_char(col_series):
        """Remove accented characters like the accented í in Santa María, for better string matches
        geopandas.geodataframe.GeoDataFrame, col_series: geopandas.geodataframe.GeoSeries
        https://stackoverflow.com/questions/33788913/pythonic-efficient-way-to-strip-whitespace-from-every-pandas-data-frame-cell-tha/33789292"""

        # Only unidecode col_series and not dataframe, does not work on unidecoding entire dataframe
        if col_series is not None:
            # column series already stripped of white space from AdminBoundaries class
            result = col_series.apply(lambda x: unidecode(x.strip()) if isinstance(x, str) else x)
            return result

    @staticmethod
    def is_valid_epsg(epsg_string):
        match = re.search('^\d{4,5}$', epsg_string)
        if match is not None:
            return True
        else:
            return False

    @staticmethod
    # User must enter range between 1 to 100 for Fuzzy Match
    def is_valid_cutoff(fuzzy_num):
        match = re.search('^([1-9][0-9]?|100)$', fuzzy_num)
        if match is not None:
            return True
        else:
            return False

    @staticmethod
    def get_file_time_stamp():
        return datetime.datetime.now().strftime('%Y_%b_%d_%Hhr_%Mmin_%Ssec')

    @staticmethod
    def create_admin_matches_shapefile(matched_records_gdf, projected_map_input, admin_choice):
        if projected_map_input is not None:
            matched_records_gdf.geometry = matched_records_gdf.geometry.to_crs(
                epsg=projected_map_input)  # works July 4 to avoid warning  Use 'GeoSeries.to_crs()'
            matched_records_gdf.geometry = matched_records_gdf.centroid
            shapefile_name = 'matches_{0}_{1}.shp'.format(
                admin_choice, DataUtility.get_file_time_stamp())
            shapefile_path = path.join(DataUtility.get_output_path(), shapefile_name)
            matched_records_gdf.to_file(driver='ESRI Shapefile', filename=shapefile_path, index=False)
            return 'Your generated admin shapefile is located at:\n{0}'.format(shapefile_path)

    @staticmethod
    def get_output_path():
        if platform.system() == 'Windows':
            dir_path = 'c:\\gis_output\\'
        elif platform.system() == 'Darwin' or platform.system() == 'Linux':
            dir_path = '/gis_output/'
        if path.isdir(dir_path):
            return dir_path
        else:
            mkdir(dir_path)
            return dir_path

    @staticmethod
    def get_file_encoding(file_path):
        """
        Detect the encoding of the given byte string from opened file, returns a string.
        :param file_path: file path to open
        :type file_path:  string
        """
        with open(file_path, 'rb') as f:
            content = f.read()
            return UnicodeDammit(content).original_encoding

    @staticmethod
    # Verify that a string can be cast to a float
    def is_float_(text):
        try:
            float(text)
            return True
        except ValueError:
            return False


class AdminBoundaries:

    def __init__(self, file_path):
        self._file_path = sanitize_filepath(file_path, platform='auto')

        if path.isfile(file_path):
            self._dataframe = geopandas.read_file(file_path)
        else:
            print(
                'The file {0} could not be located! Make sure you entered the correct file path for the admin boundaries shapefile!'.
                    format(file_path))
            exit()

    @property
    def file_path(self):
        return self._file_path

    @property
    def dataframe(self):
        return self._dataframe

    def data_column(self, col_name):
        """Return a GeoSeries of the needed admin boundaries column"""
        try:
            return self._dataframe[col_name].str.strip()
        except AttributeError:
            # print('The column {0} does not contain string values!'.format(col_name))
            return self._dataframe[col_name]

    def data_row(self, objectid):
        """Return a data row in the geodataframe based on objectid"""
        if objectid >= 0:
            return self._dataframe[objectid:objectid + 1]


class SpreadsheetData:

    def __init__(self, file_path):
        """Constructor.

        :param file_path: string for teh file path of the spreadsheet.
        """
        self._file_path = sanitize_filepath(file_path,
                                            platform='auto')

        self._western_europe_encodings = ('ascii', 'latin-1', 'utf-8', 'iso-8859-15', 'iso-8859-1')

        encoding = DataUtility.get_file_encoding(file_path)
        if encoding is not None:
            self._encoding = encoding
            print('UnicodeDammit detected encoding as {0}'.format(self._encoding))
        else:
            self._encoding = None

        if path.isfile(file_path) and file_path.lower().endswith('.csv'):
            # If detected None encoding force geopandas to read w/ 8859-1 otherwise geopandas reads w/ detected encoding
            self._dataframe = geopandas.read_file(file_path, engine='python',
                                                  encoding='iso-8859-1' if self.encoding is None else self.encoding,
                                                  errors='backslashreplace')
            # Assign encoding value ONLY ONCE to spreadsheet instance variable
            # Currently only supports western european/Latin and some Eastern European languages, uses bs4-UnicodeDammit
            # See https://stackoverflow.com/questions/8509339/what-is-the-most-common-encoding-of-each-language

            # To prevent fillna error when running fuzzy matching to a GeoDataframe created from CSV file
            self._dataframe['geometry'] = self._dataframe['geometry'].fillna(value=None)

            # Only try to decode byte characters in CSV file because CSV was opened with detected encoding.
            if self.encoding is not None:
                for col in self._dataframe.columns:
                    if col != 'geometry':
                        for val in col:
                            if isinstance(val, bytes):
                                try:
                                    val = val.decode(self.encoding)
                                except UnicodeDecodeError as ue:
                                    print('{0} error w/ {1} at {2} and ln {3}'.format(ue, self.encoding, val,
                                                                                      ue.__traceback__.tb_lineno))
                                    continue
                                except Exception as e:
                                    print('Exception {0} encountered at line {1}'.format(e, e.__traceback__.tb_lineno))
                                    continue

        elif path.isfile(file_path) and (file_path.lower().endswith('.xls') or file_path.lower().endswith('.xlsx')):
            # Must read Excel format with Pandas first and then convert to GeoDataFrame
            self._dataframe = pandas.read_excel(file_path)
            self.to_geodataframe()
            print(type(self._dataframe))

        else:
            print(
                'The file {0} could not be located! Please ensure you entered the correct spreadsheet file path!'
                    .format(file_path))
            exit()

        # Convert all column headers to lower case for easy matching by get_xy_col_locations function
        if isinstance(self._dataframe, geopandas.geodataframe.GeoDataFrame) and 'geometry' in self._dataframe.columns:
            # Convert column names to lower case and strip white space
            self._dataframe.columns = self._dataframe.columns.str.strip().str.lower()
            # Remove periods and special chars from column headers, works best on utf-8 encoded files
            self._dataframe.columns = self._dataframe.columns.str.replace(r'[.•#@&―-]', '', regex=True)
            # Set an attribute here so we know it's a Geodataframe and has a geometry column
            self.has_geom_col = 1

        # Try to load x and y or lat and log coordinates from spreadsheet into geometry values into geopandas dataframe
        if hasattr(self, 'has_geom_col'):
            if self.has_geom_col == 1:
                self.xy_to_geometry()
                print(self._dataframe['geometry'])

    # Returns a list containing x, y column numerical locations, to assign to geodatarame geometry column
    def get_xy_col_locations(self):
        """
        Get locations of x and y or longitude/latitude columns in spreadsheet.
        :return: list
        """
        xy_col_locs = []

        if 'x' in self._dataframe.columns:
            xy_col_locs.append(self._dataframe.columns.get_loc('x'))
        if 'y' in self._dataframe.columns:
            xy_col_locs.append(self._dataframe.columns.get_loc('y'))

        if len(xy_col_locs) == 2:
            return xy_col_locs

        if 'longitude' in self._dataframe.columns:
            xy_col_locs.append(self._dataframe.columns.get_loc('longitude'))
        if 'latitude' in self._dataframe.columns:
            xy_col_locs.append(self._dataframe.columns.get_loc('latitude'))
        if len(xy_col_locs) == 2:
            return xy_col_locs

        if 'lon' in self._dataframe.columns:
            xy_col_locs.append(self._dataframe.columns.get_loc('lon'))
        elif 'long' in self._dataframe.columns:
            xy_col_locs.append(self._dataframe.columns.get_loc('long'))
        if 'lat' in self._dataframe.columns:
            xy_col_locs.append(self._dataframe.columns.get_loc('lat'))
        if len(xy_col_locs) == 2:
            return xy_col_locs

    @property
    def file_path(self):
        """
        Return file path.
        """
        return self._file_path

    @property
    def data_frame(self):
        """
        Return dataframe.
        """
        return self._dataframe

    @data_frame.setter
    def data_frame(self, value):
        """Assign new dataframe to the dataframe"""
        self._data_frame = value

    @property
    def columns(self):
        """Return columns."""
        return self._dataframe.columns

    @property
    def encoding(self):
        """Return encoding."""
        return self._encoding

    @property
    def western_europe_encodings(self):
        """Return the encodings we will accept as Western European encodings, for unidecode to process."""
        return self._western_europe_encodings

    def to_pandas_dataframe(self):
        """
        Convert Geodataframe to Pandas dataframe, is a void type function, does not return any value.
        """
        if hasattr(self, 'has_geom_col'):
            if self.has_geom_col == 1:
                temp_pd_data_frame = self._dataframe
                # Must use del as temp_pd_data_frame.drop('geometry', 1) doesn't work
                del temp_pd_data_frame['geometry']
                self._dataframe = pandas.DataFrame(temp_pd_data_frame)

    # is a void type function, does not return anything
    def to_geodataframe(self):
        """
        Change a Pandas dataframe to Geodataframe. This is a void type function, does not return any value.
        """
        if isinstance(self._dataframe, pandas.core.frame.DataFrame):

            temp_array = numpy.full(len(self._dataframe), fill_value=-1.0)
            gdf = geopandas.GeoDataFrame(self._dataframe,
                                         geometry=geopandas.points_from_xy(x=temp_array, y=temp_array))
            self._dataframe = gdf

            # if has x y in columns etc.
            if 'x' and 'y' or 'lat' and 'lon' or 'lat' and 'long' \
                    or 'latitude' and 'longitude' in self._dataframe.columns:
                self.xy_to_geometry()

    def xy_to_geometry(self):
        """
        Assigns x/y lat/long column values in spreadsheet to geometry column. This is Void type function.
        """
        if self._dataframe is not None:
            if 'x' and 'y' or 'lat' and 'lon' or 'lat' and 'long' \
                    or 'latitude' and 'longitude' in self._dataframe.columns:
                xy_col_locs = self.get_xy_col_locations()
                if xy_col_locs is not None and len(xy_col_locs) == 2:
                    # Cast as string type first
                    temp_x_coords = self._dataframe.iloc[:, xy_col_locs[0]].astype(str)
                    temp_y_coords = self._dataframe.iloc[:, xy_col_locs[1]].astype(str)

                    # Make sure they're all float type before creating Point geometry values
                    # Invalid Coordinates are flagged as x = -1.0 and y = -1.0
                    temp_x_coords = temp_x_coords.apply(lambda x: float(x) if DataUtility.is_float_(x) else -1.0)
                    temp_y_coords = temp_y_coords.apply(lambda y: float(y) if DataUtility.is_float_(y) else -1.0)
                    self._dataframe['geometry'] = geopandas.points_from_xy(x=temp_x_coords, y=temp_y_coords)
                    print(
                        'X and Y or Lat and Long values were detected in the spreadsheet file and assigned to Geometry column!')
            else:
                print('X and Y coordinates were not detected in {0}!'.format(self._file_path))


class Report:
    """Creates an Excel file report to show spreadsheet matched to admin boundaries shapefile"""

    def __init__(self, spreadsheet_dataframe, admin_dataframe):
        if not isinstance(spreadsheet_dataframe, geopandas.geodataframe.GeoDataFrame):
            raise TypeError('The spreadsheet dataframe must be of the GeoDataFrame type!!')
        if not isinstance(admin_dataframe, geopandas.geodataframe.GeoDataFrame):
            raise TypeError('Administrative boundary dataframe must be of the GeoDataFrame type!!')
        if isinstance(spreadsheet_dataframe, geopandas.geodataframe.GeoDataFrame) and \
                isinstance(admin_dataframe, geopandas.geodataframe.GeoDataFrame):
            self._spreadsheet_dataframe = spreadsheet_dataframe
            self._admin_dataframe = admin_dataframe
            self._joined_dataframe = None

    def join_dataframes(self):
        # To prevent ValueError where columns overlap
        del self._spreadsheet_dataframe['geometry']
        self._joined_dataframe = self._spreadsheet_dataframe.join(self._admin_dataframe.set_index('Index'), on='Index',
                                                                  how='inner')

    def save_report(self):
        # Message strings are returned so that it can be passed to GUI window, or to console print statement
        if self._joined_dataframe is not None:
            out_file = path.join(DataUtility.get_output_path(),
                                 'match_excel_report_{0}.xlsx'.format(DataUtility.get_file_time_stamp()))
            self._joined_dataframe.to_excel(out_file)
            return 'The report of spreadsheet records matched' \
                   '\nto the admin boundaries shapefile data has been saved at: {0}'.format(out_file)
        else:
            return 'Unable to save the report of spreadsheet\nrecords matched to the admin boundaries shapefile data' \
                   '\nbecause the joined dataframe is None.'


class MatchedData:
    """
    This class represents any matches between the spreadsheet data and the admin boundaries data
    """

    def __init__(self, spreadsheet_data, adm_boundaries):
        """Constructor.
        :param spreadsheet_data: string for spreadsheet data file
        :param adm_boundaries: string for the admin boundaries shapefile
        """
        self._admin_choice = None
        self._spreadsheet_data = spreadsheet_data
        self._adm_boundaries = adm_boundaries

        # Stores matched data according to spreadsheet row, value is namedtuple('row_data', ['shp_data', 'sheet_data'])
        self._matched_data_dict = OrderedDict()
        self._unmatched_data_dict = OrderedDict()

    @property
    def spreadsheet_data(self):
        return self._spreadsheet_data

    @property
    def adm_boundaries(self):
        return self._adm_boundaries

    @property
    def matched_spreadsheet_with_score_gdf(self):
        return self._matched_spreadsheet_with_score_gdf

    @property
    def matched_admin_dict(self):
        return self._matched_admin_dict

    @property
    def matched_data_dict(self):
        return self._matched_data_dict

    @property
    def unmatched_data_dict(self):
        return self._unmatched_data_dict

    @property
    def admin_choice(self):
        return self._admin_choice

    @admin_choice.setter
    def admin_choice(self, value):
        self._admin_choice = value

    def get_admin_choices(self):
        """
        Get Admin choices
        :return: dictionary with keys storing column indices, and values storing admin boundary column names
        """
        columns_dict = dict(zip([str(self._adm_boundaries.dataframe.columns.get_loc(c)) for c in
                                 self._adm_boundaries.dataframe.columns], self._adm_boundaries.dataframe.columns))
        return columns_dict

    def selected_admin_choice(self):
        """
        Attribute to indicate user selected Admin Choice, possibly only use in console version?, Not Needed for GUI.
        """
        self.selected_admin_choice = 1

    def user_proceed_match(self):
        """
        Attribute showing that user has indicated to proceed with match
        """
        self.user_proceed_match = 1

    def get_spreadsheet_report_dataframe(self):
        """
        Dataframe used for generating Excel file match report, passed to the Report class
        """
        if len(self._matched_data_dict) > 0:
            '''old version temp_list = [v[1] for k,v in self._matched_data_dict.items() ]
            new ver w/ NamedTuple temp_list = [val.sheet_data for key, val in self._matched_data_dict.items()]
            debug print('Temp_List for MatcheData.get_spreadsheet_report_dataframe {0}'.format(temp_list))'''
            temp_list = [val.sheet_data for key, val in self._matched_data_dict.items()]
            temp_df = pandas.concat(temp_list, axis=1).transpose()
            temp_geom_array = numpy.full(len(temp_df), fill_value=-1.0)
            temp_gdf = geopandas.GeoDataFrame(temp_df, geometry=geopandas.points_from_xy(
                x=temp_geom_array, y=temp_geom_array))
        return temp_gdf

    def array_to_series(self, row, score):
        """Convert filtered ndarray from dataframe to Pandas series, inserts match score in the returned Pandas series.
        :param row: ndarray of the matched row.
        :param score: integer for the match score.
        :return: Pandas series for the matched row.
        """
        row_series = pandas.Series(row)
        row_series.index = row._fields
        row_series['Match_Score'] = score
        return row_series

    # This match is always run and does strict text matching
    def run_strict_match(self, **kwargs):
        """Always run the strict match first.
        :param **kwargs: dictionary keyword argument. Only valid keyword argument is from_right_col: 1.
        """
        # Loop through each column, Pandas' first col value starts at index 1
        col_size = len(self._spreadsheet_data.data_frame.columns)
        print('kwargs passed to run_match function: {0}'.format(kwargs.keys()))
        # Only start searching from most right sided column if user wants it.
        if kwargs.get('from_right_col') == 1:
            for row in self._spreadsheet_data.data_frame.itertuples():
                # Ensure insertion of one record per row
                insertions = 0
                print('MatchData - checking for matches between spreadsheet and admin boundaries shapefile...')
                for i in reversed(range(1, col_size)):
                    if DataUtility.is_string_match(row[i], self._adm_boundaries.data_column(self._admin_choice),
                                                   self.spreadsheet_data):
                        data_row = DataUtility.filter_row(self._adm_boundaries.dataframe, self._admin_choice, row[i],
                                                          self.spreadsheet_data)
                        # Prevent inserting more than one match from multiple columns
                        if insertions == 0:
                            row_data = namedtuple('row_data', ['shp_data', 'sheet_data'])
                            row_data.shp_data = data_row
                            # self._matched_admin_dict[row.Index] = data_row
                            # Track matches in spreadsheet file
                            row_data.sheet_data = self.array_to_series(row, 100)
                            # Add to the data dict
                            self._matched_data_dict[row.Index] = row_data
                            insertions += 1
                            print('Added Spreadsheet row number {0} to matches!'.format(row.Index))
                        continue
                if row.Index not in self._matched_data_dict.keys():
                    self._unmatched_data_dict[row.Index] = row

        # Default is searching from left to right columns in spreadsheet
        else:
            for row in self._spreadsheet_data.data_frame.itertuples():
                insertions = 0
                print('MatchData - checking for matches between spreadsheet and admin boundaries shapefile...')
                for i in range(1, col_size):
                    if DataUtility.is_string_match(row[i], self._adm_boundaries.data_column(self._admin_choice),
                                                   self.spreadsheet_data):
                        data_row = DataUtility.filter_row(self._adm_boundaries.dataframe, self._admin_choice, row[i],
                                                          self.spreadsheet_data)
                        if insertions == 0:
                            # Info for shapefile
                            row_data = namedtuple('row_data', ['shp_data', 'sheet_data'])
                            row_data.shp_data = data_row
                            # Track matches in spreadsheet file
                            # Save rows as Pandas series along with match score
                            row_data.sheet_data = self.array_to_series(row, 100)
                            # Add to the data dict
                            self._matched_data_dict[row.Index] = row_data
                            insertions += 1
                            print('Added Spreadsheet row number {0} to matches!'.format(row.Index))
                        continue

                if row.Index not in self._matched_data_dict.keys():
                    self._unmatched_data_dict[row.Index] = row

    def run_fuzzy_match(self, min_score, **kwargs):
        """Fuzzy match function, only executed when user selects fuzzy matching.
        :param **kwargs: dictionary keyword argument. Only valid keyword argument is from_right_col: 1.
        """
        # Loop through each column, Pandas' first col value starts at index 1, Aug. 21 I need index 0 for the count idx!
        col_size = len(self._spreadsheet_data.data_frame.columns)
        print('Fuzzy match running on file type: {0}. '.format(self.spreadsheet_data))
        print('Unmatched data dict used for fuzzy matching: {0}'.format(self._unmatched_data_dict))
        # Using Pandas iloc causes
        # SettingWithCopyWarning: self._spreadsheet_data.data_frame.iloc[list(self._unmatched_data_dict.keys()), :]
        # Better to do df.loc[[7,8,9]] than to do df.iloc[[7,8,9],:]
        fuzzy_spreadsheet_df = self._spreadsheet_data.data_frame.loc[list(self._unmatched_data_dict.keys())]

        for col in fuzzy_spreadsheet_df.columns:
            if col != 'geometry':
                # Strip out NaN cell values as they throw off the fuzzy matching, cast as string for the fuzzy match
                fuzzy_spreadsheet_df[col] = fuzzy_spreadsheet_df[col].fillna('###')
                fuzzy_spreadsheet_df[col] = fuzzy_spreadsheet_df[col].astype(str)

        temp_adm_boundaries_df = self._adm_boundaries.dataframe
        temp_adm_boundaries_list = temp_adm_boundaries_df[self._admin_choice].astype(str).values.tolist()

        # User did not provide priority right column option, so we do the default search order from left to right
        if kwargs.get('from_right_col') is None:
            for row in fuzzy_spreadsheet_df.itertuples():
                insertions = 0
                print(
                    'Fuzzy Match from the Left - checking for matches between spreadsheet and admin boundaries shapefile...')
                for i in range(1, col_size):
                    best_match = self.fuzzy_match_text(row[i], temp_adm_boundaries_list, int(min_score))
                    if best_match is not None:
                        data_row = DataUtility.filter_row(self._adm_boundaries.dataframe, self._admin_choice,
                                                          best_match[0], self.spreadsheet_data)  # row[i] in spreadsheet
                        if insertions == 0:
                            # shapefile info
                            row_data = namedtuple('row_data', ['shp_data', 'sheet_data'])
                            row_data.shp_data = data_row
                            row_data.sheet_data = self.array_to_series(row, best_match[1])
                            # Add to the data dict
                            self._matched_data_dict[row.Index] = row_data
                            insertions += 1
                            print('Added FUZZY MATCHED Spreadsheet row number {0} to matches!'.format(row.Index))
                        continue

        # Only search from most right sided columns if user wants it.
        elif kwargs.get('from_right_col') == 1:
            for row in fuzzy_spreadsheet_df.itertuples():
                insertions = 0
                print(
                    'Fuzzy Match from the RIGHT - checking for matches between spreadsheet and admin boundaries shapefile...')
                for i in reversed(range(1, col_size)):
                    best_match = self.fuzzy_match_text(row[i], temp_adm_boundaries_list, int(min_score))
                    if best_match is not None:
                        data_row = DataUtility.filter_row(self._adm_boundaries.dataframe, self._admin_choice,
                                                          best_match[0], self.spreadsheet_data)  # row[i] in spreadsheet
                        if insertions == 0:
                            # Info for shpfile
                            row_data = namedtuple('row_data', ['shp_data', 'sheet_data'])
                            row_data.shp_data = data_row
                            row_data.sheet_data = self.array_to_series(row, best_match[1])

                            self._matched_data_dict[row.Index] = row_data
                            insertions += 1
                            print('Added FUZZY MATCHED Spreadsheet row number {0} to matches!'.format(row.Index))
                        continue

    def fuzzy_match_text(self, text_to_match, options, min_score):
        """
        Fuzzy match the text provided.
        Will try to match min_score and above https://github.com/seatgeek/fuzzywuzzy/blob/master/fuzzywuzzy/process.py
        :param text_to_match: string of the text.
        :param options: A list or dictionary of choices in the fuzzy match.
        :param min_score: Integer, Optional argument for score threshold.
        :return: Tuple containing a single match and its score, or None.
        """
        #
        result = process.extractOne(query=text_to_match, choices=options, scorer=fuzz.WRatio,
                                    score_cutoff=min_score)
        if isinstance(result, tuple) and len(result) == 2:
            # Must cast as tuple for the result to be readable by other functions
            return tuple(result)
        else:
            return None


# The functions below are used by the console version of the application
def prompt_for_admin_area_console(md):
    # dataframe is geopandas.GeoDataFrame
    # Create a choice of admin areas to select
    columns_dict = md.get_admin_choices()
    print('\nNow that you have selected an admin boundary shapefile/polygon.')
    print('Please select the Field with the region names with which you want to match to the spreadsheet data:')
    for key, val in columns_dict.items():
        print('Press {0} to select {1}'.format(key, val))
    print('HINT: You generally don\'t want to select shapefile columns with names like OBJECTID_1 and geometry as'
          ' they are not admin areas!')
    print(
        'Please check online GIS resources or shapefile\'s metadata to confirm which shapefile field name to select.')
    return columns_dict


def print_console_help():
    print('Instructions are available by typing: python match_admin_boundaries_core.py --help')
    print('Example with regular arguments entered:')
    print(
        'python match_admin_boundaries_core.py --spreadsheet_file "c:\\temp\\data.xlsx" '
        '--admin_boundaries_file "c:\\gisdata\\hnd_admbnda_adm3_sinit_20161005.shp" --match_type regular')
    print('Example with abbreviated arguments entered:')
    print(
        'python match_admin_boundaries_core.py -s "c:\\temp\\data.xlsx" '
        '-a "c:\\gisdata\\hnd_admbnda_adm3_sinit_20161005.shp" -m fuzzy')


def run_console_match(arg_val, md):
    if arg_val.lower().strip() == 'fuzzy':
        p = PromptMessages()
        p.argument = 'hit enter key'
        print(p.fuzzy_caption)
        fuzzy_input = str(input('Enter fuzzy match cutoff score between 1 and 99. --> '))
        if DataUtility.is_valid_cutoff(fuzzy_input):
            print('Fuzzy cut-off score {0} entered, proceeding with fuzzy match.'.format(fuzzy_input))
            process_column_priority(arg_val, md, fuzzy_input=fuzzy_input)
        else:
            print('{0} is an invalid cutoff score. Please rerun this program from the beginning!'.format(fuzzy_input))
            # Exit code 1 Invalid cutoff score
            exit(1)
    elif arg_val.lower().strip() == 'regular':
        print('Proceeding to do regular match')
        process_column_priority(arg_val, md)


def process_column_priority(match_arg_val, md, **kwargs):
    print(
        'If you want to choose the columns on the right side of spreadsheet, type \'priority_right\' & hit enter key.')
    print('Otherwise type any key and hit Enter.')
    col_pri_input = str(input('Enter column priority and hit Enter key. --> ')).lower().strip()

    try:
        if match_arg_val == 'regular':
            if col_pri_input == 'priority_right':
                md.run_strict_match(from_right_col=1)
            # Any other key(s) were entered.
            else:
                md.run_strict_match()
                print('Length was {0}'.format(md.matched_data_dict))

        elif match_arg_val == 'fuzzy':
            if col_pri_input == 'priority_right':
                md.run_strict_match(from_right_col=1)
                md.run_fuzzy_match(kwargs.get('fuzzy_input'), from_right_col=1)
            # Any other key(s) were entered.
            else:
                md.run_strict_match()
                md.run_fuzzy_match(kwargs.get('fuzzy_input'))

    except KeyError as e:
        print('You need to enter a column priority. Enter the word regular or priority_right and hit Enter key!')
    except Exception as e:
        print('Exception {0} encountered at line {1}'.format(e, e.__traceback__.tb_lineno))


def main():
    try:
        parser = ArgumentParser(description=__doc__, formatter_class=RawDescriptionHelpFormatter)
        parser.add_argument('-s',
                            '--spreadsheet_file',
                            type=str,
                            help='File location of spreadsheet file. It can be CSV (.csv), or Excel (.xlsx or .xls) format')
        parser.add_argument('-a',
                            '--admin_boundaries_file',
                            type=str,
                            help='File location of your admin boundaries shapefile. It should be a shapefile-ends in .shp.')
        parser.add_argument('-m',
                            '--match_type',
                            type=str,
                            help='Choose regular match or fuzzy match.')
        args = parser.parse_args()

        if not (args.spreadsheet_file and args.admin_boundaries_file and args.match_type):
            print(
                '\nYou need to provide 3 arguments: a file location for the spreadsheet file, a file location for the '
                'admin boundaries shapefile, and choose a match type.')
            print_console_help()
            return

        else:
            md = MatchedData(SpreadsheetData(args.spreadsheet_file), AdminBoundaries(args.admin_boundaries_file))
            continue_admin_prompt = True
            while continue_admin_prompt:
                admin_boundaries_dict = prompt_for_admin_area_console(md)
                admin_input = str(input('Please enter your choice of administrative area. --> ')).strip()
                if admin_input in admin_boundaries_dict.keys():
                    admin_choice = admin_boundaries_dict[admin_input]
                    continue_admin_prompt = False
                    md.admin_choice = admin_choice
                    md.user_proceed_match()
                    if hasattr(md, 'user_proceed_match') and md.admin_choice is not None:
                        if md.user_proceed_match == 1:
                            run_console_match(args.match_type, md)
                elif admin_input not in admin_boundaries_dict.keys():
                    print('\nYou entered an Invalid choice for administrative area. Please try again!\n')
                    print('\nIf you want to stop this program, type the x key and hit Enter to stop this program.')
                elif admin_input.lower().strip() == 'x':
                    exit()

        args = parser.parse_args()
        print(md.spreadsheet_data)
        print(md.adm_boundaries)

        shp_file = AdminBoundaries(args.admin_boundaries_file)

        # Set the row number to match the csv/Excel row numbering
        md.spreadsheet_data.data_frame.index = md.spreadsheet_data.data_frame.index + 2

        print('Simplified spreadsheet matches as matched admin dict!!')
        [print('Key: {0} Value: {1}'.format(k, v)) for k, v in md.matched_data_dict.items()]
        print('{0}'.format(len(md.matched_data_dict)))
        print('\n')

        if len(md.matched_data_dict) > 0:
            print('There are {0} records matched to the admin boundaries shapefile'.format(len(md.matched_data_dict)))
            print('{0} spreadsheet records matched to the admin boundaries shapefile\nout of a total of '
                  '{1} spreadsheet records'.format(len(md.matched_data_dict),
                                                   len(md.spreadsheet_data.data_frame.index)))

            print('Would you like to create a Excel report file to show the matches?')
            print('\nEnter Y to create Excel report file, or press any other key skip this step.')
            excel_report_input = str(input('Create Excel file report of matches. --> ')).lower().strip()
            if excel_report_input == 'y' or excel_report_input == 'yes':
                temp_df = md.get_spreadsheet_report_dataframe()
                report_df = geopandas.GeoDataFrame(
                    data=temp_df, crs="EPSG:4326",
                    geometry=temp_df['geometry'])  # data=[itesm for item in md.matched_data_dict.values()]
                report_df.set_index('Index')

                admin_shapefile_df = geopandas.GeoDataFrame(
                    # Is no longer data=[val[0] for val in md.matched_data_dict.values()]
                    data=[val.shp_data for val in md.matched_data_dict.values()], crs="EPSG:4326",
                    columns=shp_file.dataframe.columns)

                admin_shapefile_df['Index'] = md.matched_data_dict.keys()
                admin_shapefile_df.set_index('Index')
                report = Report(report_df, admin_shapefile_df)
                report.join_dataframes()
                print(report.save_report())

            print('\nWould you like to create a shapefile to show the matches on a map?')
            print('\nEnter Y to create the shapefile, or press any other key finish this program.')
            shpfile_prompt_input = str(input('Create shapefile to show the matches. --> ')).lower().strip()
            if shpfile_prompt_input == 'y' or shpfile_prompt_input == 'yes':
                valid_epsg_input = False
                while not valid_epsg_input:
                    try:
                        p = PromptMessages()
                        p.argument = 'hit enter key'
                        print(p.epsg_caption)
                        epsg_input = str(
                            input(
                                'Enter the 4 or 5 digit EPSG code that you found from one of the above websites. --> ')).lower().strip()
                        epsg_match = DataUtility.is_valid_epsg(epsg_input)

                        if epsg_input == 'x' or epsg_input == 'exit':
                            exit(0)
                        elif not epsg_match:
                            raise ValueError
                        elif epsg_match:

                            matched_admin_list = [val.shp_data for val in md.matched_data_dict.values()]

                            if 'geometry' in shp_file.dataframe.columns:
                                matched_geom_col_loc = shp_file.dataframe.columns.get_loc('geometry')

                            # Create geodataframe and output to shapefile
                            matched_records_gdf = geopandas.GeoDataFrame(data=matched_admin_list,
                                                                         columns=shp_file.dataframe.columns,
                                                                         crs="EPSG:4326",
                                                                         geometry=numpy.asarray(list(
                                                                             [row[matched_geom_col_loc] for row in
                                                                              matched_admin_list])))
                            print(DataUtility.create_admin_matches_shapefile(matched_records_gdf,
                                                                             epsg_input,
                                                                             md.admin_choice))
                            valid_epsg_input = True
                    except ValueError:
                        print(
                            '\nYou entered {0} which is an invalid epsg code. EPSG Codes must be 4 or 5 numerical digits '
                            'only! Please try again.'.format(epsg_input))
                        print('Enter x if you wish to exit this program.\r\n')
                    except Exception as e:
                        print('Exception {0} occurred.'.format(e))
        elif len(md.matched_data_dict) == 0:
            print('No matches were found between the spreadsheet file and the admin boundaries shapefile!')
            print('Please try again!!')
    except Exception as e:
        print('Exception {0} at line {1}'.format(e, e.__traceback__.tb_lineno))


if __name__ == "__main__":
    main()
