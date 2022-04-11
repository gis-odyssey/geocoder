import geopandas
import numpy
from match_admin_boundaries_core import SpreadsheetData, AdminBoundaries, MatchedData, DataUtility, Report, \
    PromptMessages
import wx
from wx.lib import sized_controls
import sys

class Frame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, parent=None,
                          title='App to Match Admin Boundaries', size = ( 800, 800 ))

        self.panel = wx.Panel(self, -1 )
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Title
        self.centred_text = wx.StaticText(self.panel, label="App to Match Admin Boundaries")
        self.main_sizer.Add(self.centred_text, 0, wx.ALIGN_CENTRE | wx.ALL, 3)

        # Grids
        self.content_sizer = wx.BoxSizer(wx.VERTICAL)
        self.grid_1 = wx.GridSizer(1, 3, 2, 2) #GridSizer(rows, cols, vgap, hgap)
        self.spreadsheet_btn = wx.Button(self.panel, label='Choose Spreadsheet File (.XLS, .XLSX, or .CSV)') #, pos=(30, 30)
        self.spreadsheet_btn.Bind(wx.EVT_BUTTON, self.on_open_spreadsheet)
        self.admin_btn = wx.Button(self.panel, label='Choose Administrative Boundary Shapefile')
        self.admin_btn.Bind(wx.EVT_BUTTON, self.on_open_shapefile)
        self.match_btn = wx.Button(self.panel, label='Run the Matching Process')
        self.match_btn.Bind(wx.EVT_BUTTON, self.on_press_match_btn)
        self.grid_1.AddMany([self.spreadsheet_btn, self.admin_btn, self.match_btn])
        # is  35 pixel border Can't use wx.ALIGN_CENTER with wx.ALL and wx.EXPAND
        self.content_sizer.Add(self.grid_1, -1,  wx.ALL | wx.EXPAND, 35)

        # Frame to show console output
        # -1 is not stretchable when maximized window size=(400, 150)
        console_text = wx.TextCtrl(self.panel, -1, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)
        redir = RedirectText(console_text)
        sys.stdout = redir
        self.content_sizer.Add(console_text, -1,  wx.EXPAND, 8)  # 8 pixels border/padding

        # Declare spreadsheet and shapefile variables
        self.spreadsheet = None
        self.shapefile = None

        self.nb = wx.Notebook(self.panel)
        self.content_sizer.Add(self.nb, -1,  wx.EXPAND, 12)
        self.main_sizer.Add(self.content_sizer, -1, wx.EXPAND)
        self.panel.SetSizer(self.main_sizer)
        self.Show()

    def on_open_spreadsheet(self, event):

        # Ask the user what new file to open
        with wx.FileDialog(self, "Open Spreadsheet file", wildcard = "Microsoft Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|"
                                                                     "CSV files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            # Proceed loading the file chosen by the user
            pathname = fileDialog.GetPath()
            try:
                with open(pathname, 'r') as file:
                    print('Spreadsheet {0}'.format(pathname))
                    #print(file.read())
                    self.spreadsheet = SpreadsheetData(pathname)

                    df_tab = PreviewTable(self.nb, self.spreadsheet.data_frame) #DataframePanel(nb, df, self.status_bar_callback)
                    #self.tab_num += 1
                    self.nb.AddPage(df_tab, "  Tab %s" % pathname)

            except IOError:
                wx.LogError("Cannot open file '%s'." % file)

    def on_open_shapefile(self, event):
        with wx.FileDialog(self, "Open Admin Boundary file", wildcard = "Shapefiles (*.shp)|*.shp",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            pathname = fileDialog.GetPath()
            try:
                with open(pathname, 'r') as file:
                    print('Shapefile {0}'.format(pathname))

                    self.shapefile = AdminBoundaries(pathname)
                    df_tab = PreviewTable(self.nb, self.shapefile.dataframe)
                    #self.tab_num += 1
                    self.nb.AddPage(df_tab, "  Tab: %s" % pathname)
            except IOError:
                wx.LogError("Cannot open file '%s'." % file)

    def on_press_match_btn(self, event):
        if self.spreadsheet is not None and self.shapefile is not None:

            md = MatchedData(self.spreadsheet, self.shapefile)

            #admin_choice = None

            rad_box_text = '\nNow that you selected your boundary polygon shapefile, please select the field ' \
                           '\nthat has the the region names that you are trying to match.' \
                           '\nChoose the field from the Radio Button choices shown on the left side of this window.' \
                           '\nHINT: You generally don\'t want to select fields with names like OBJECTID_1' \
                           ' or geometry.' \
                           '\nCheck online GIS resources or the shapefile\'s metadata to verify the ' \
                           'field you should choose.'

            admin_col_dict = md.get_admin_choices()
            rbd = RadioBoxDialog(None, rad_box_text, list(admin_col_dict.values()))

            rbd_obj = rbd.ShowModal()

            if hasattr(rbd, 'radio_box_pressed_ok_btn'):

                if hasattr(rbd, 'fuzzy_match'):
                    p = PromptMessages()
                    p.argument = 'click OK button'
                    fuzzy_dlg = FuzzyDialog(None, 'Fuzzy Match Selected', p.fuzzy_caption)
                    if fuzzy_dlg.ShowModal() == wx.ID_OK:
                        if DataUtility.is_valid_cutoff(fuzzy_dlg.input.GetValue()):
                            md.admin_choice = rbd.radbox_admin_choice
                            md.user_proceed_match()

                            # User chose right handed column as priority
                            if rbd.col_rad_box_choice == 'Prioritize Right Column':
                                md.run_strict_match(from_right_col=1)
                                md.run_fuzzy_match(fuzzy_dlg.GetValue(), from_right_col=1)

                            # Use Default order of matching
                            else:
                                md.run_strict_match()
                                md.run_fuzzy_match(fuzzy_dlg.GetValue())

                        else:
                            wrong_cutoff_dlg = wx.MessageDialog(None,
                                                                'Please enter a valid number that is between 1 and 99',
                                                                'Incorrect Cut-off Score entered', wx.OK)
                            wrong_cutoff_dlg.ShowModal()
                            wrong_cutoff_dlg.Destroy()

                else:
                    # Do not call setter as self.md.admin_choice(value)
                    md.admin_choice = rbd.radbox_admin_choice
                    md.user_proceed_match()

                    if rbd.col_rad_box_choice == 'Prioritize Right Column':
                        md.run_strict_match(from_right_col=1)
                    else:
                        md.run_strict_match()

            if hasattr(md, 'user_proceed_match') and md.admin_choice is not None:

                if len(md.matched_data_dict) > 0:
                    print('There are {0} records matched to the admin boundaries shapefile'.format(
                        len(md.matched_data_dict)))
                    yes_no_dlg = wx.MessageDialog(None,
                                                  '{0} spreadsheet records matched to the admin boundaries shapefile'
                                                  '\nout of a total of {1} spreadsheet records'
                                                  '\nContinue to next Step? Choose Yes or No.'.format(
                                                      len(md.matched_data_dict),
                                                      len(md.spreadsheet_data.data_frame.index)),
                                                  "Matches Found", wx.YES_NO)
                    reply = yes_no_dlg.ShowModal()
                    # Now destroy dialog to prevent needing double press of button
                    yes_no_dlg.Destroy()
                    # 2 is YES, 8 is NO  wx.YES/NO don't work, must use ID_YES etc.

                    if reply == wx.ID_YES:

                        report_dlg = wx.MessageDialog(None,
                                                      'Create an Excel report that shows the spreadsheet data'
                                                      '\nmatched to the admin boundaries shapefile?'
                                                      '\nClick OK to create the Excel report file',
                                                      'Create Excel Report of Matches', wx.YES_NO | wx.CANCEL)

                        report_dlg.SetYesNoCancelLabels(wx.ID_OK, "Skip This Step and Create Shapefile", wx.ID_CANCEL)
                        report_response = report_dlg.ShowModal()
                        report_dlg.Destroy()

                        print(report_response)

                        if report_response == wx.ID_YES:
                            try:
                                temp_df = md.get_spreadsheet_report_dataframe()
                                report_df = geopandas.GeoDataFrame(
                                    data=temp_df, crs="EPSG:4326", geometry=temp_df['geometry'])
                                report_df.set_index('Index')

                                # Add the spatial info from data_dict
                                admin_shapefile_df = geopandas.GeoDataFrame(
                                    # Is no longer data=[val[0] for val in md.matched_data_dict.values()]
                                    data=[val.shp_data for val in md.matched_data_dict.values()], crs="EPSG:4326",
                                    columns=self.shapefile.dataframe.columns)
                                admin_shapefile_df['Index'] = md.matched_data_dict.keys()
                                admin_shapefile_df.set_index('Index')
                                report = Report(report_df, admin_shapefile_df)
                                report.join_dataframes()
                                excel_msg = report.save_report()

                                excel_report_dlg = wx.MessageDialog(None,
                                                                    excel_msg,
                                                                    'Excel file report of matches created', wx.OK)
                                excel_report_dlg.ShowModal()
                                excel_report_dlg.Destroy()

                                self.prompt_create_admin_shapefile(md)

                            except Exception as e:
                                print(
                                    'Exception {0} occurred while trying to create the Excel report of the matches'.format(
                                        e))

                        elif report_response == wx.ID_NO:
                            self.prompt_create_admin_shapefile(md)

                    else:
                        print('No Selected!')

                elif len(md.matched_data_dict) == 0 or md.matched_data_dict is None:
                    print('No matches found in MatchData matched_data_dict, {0} matches.'.format(len(md.matched_data_dict)))
                    no_match_dlg = wx.MessageDialog(None,
                                                    'No spreadsheet matches were found in the shapefile.'
                                                    '\nTry selecting another spreadsheet/admin boundaries shapefile.',
                                                    "No Matches", wx.OK)
                    no_match_dlg.ShowModal()
                    no_match_dlg.Destroy()

    # md is the matched_data
    def prompt_create_admin_shapefile(self, md):

        p = PromptMessages()
        p.argument = 'click OK button'

        epsg_dlg = EPSGDialog(None, 'Enter EPSG Code', p.epsg_caption)

        if epsg_dlg.ShowModal() == wx.ID_OK:

            if DataUtility.is_valid_epsg(epsg_dlg.GetValue()):
                try:
                    matched_admin_list = [val.shp_data for val in md.matched_data_dict.values()]

                    if 'geometry' in self.shapefile.dataframe.columns:
                        matched_geom_col_loc = self.shapefile.dataframe.columns.get_loc('geometry')

                    # Create geodataframe and output to shapefile
                    matched_records_gdf = geopandas.GeoDataFrame(data=matched_admin_list,
                                                                 columns=self.shapefile.dataframe.columns,
                                                                 crs="EPSG:4326",
                                                                 geometry=numpy.asarray(list(
                                                                     [row[matched_geom_col_loc] for row in
                                                                      matched_admin_list])))
                    shapefile_msg = DataUtility.create_admin_matches_shapefile(matched_records_gdf, epsg_dlg.GetValue(),
                                                                               md.admin_choice)
                    shapefile_dlg = wx.MessageDialog(None,
                                                     shapefile_msg,
                                                     'Matched data Shapefile created', wx.OK)
                    shapefile_dlg.ShowModal()
                    shapefile_dlg.Destroy()

                except Exception as e:
                    print('Exception {0} occurred while trying to save the matches shapefile'.format(e))

            else:
                incorrect_epsg_dialog = wx.MessageDialog(None,
                                                         'No shapefile was created because the wrong epsg code was entered.' \
                                                         '\nPlease enter an epsg code that contains 4 or 5 numerical digits.',
                                                         'Incorrect EPSG Code', wx.OK)
                incorrect_epsg_dialog.ShowModal()
                incorrect_epsg_dialog.Destroy()


class RadioBoxDialog(sized_controls.SizedDialog):

    def __init__(self, parent, rad_box_text, radio_list):
        """constructor"""
        wx.Dialog.__init__(self, parent, size=(875, 730))

        self.panel = wx.Panel(self)
        self.instructions = wx.StaticText(self.panel, label=rad_box_text, pos=(10, 10))

        self.radio_box = wx.RadioBox(self.panel, label='Admin Boundaries', choices=radio_list, majorDimension=0,
                                style=wx.RA_SPECIFY_ROWS, pos=(5,118))
        self.radio_box.Bind(wx.EVT_RADIOBOX, self.on_radio_group)

        self.radbox_admin_choice = self.radio_box.GetStringSelection()

        col_select_text = 'Choose ONE Column Priority for Searching:\nIf you\'re not sure,'\
        ' choose Regular.'
        self.column_instructions = wx.StaticText(self.panel, label=col_select_text, pos=(200, 140))

        column_list = ['Regular', 'Prioritize Right Column']
        self.col_rad_box = wx.RadioBox(self.panel, label='Column Priority', pos=(200, 185),
                                       choices=column_list, majorDimension=0, style=wx.RA_SPECIFY_ROWS)
        self.col_rad_box.Bind(wx.EVT_RADIOBOX, self.on_col_radio_box)
        self.col_rad_box_choice = self.col_rad_box.GetStringSelection()
        print('Radio box DEFAULT admin choice: {0}'.format(self.radbox_admin_choice))
        print('Column priority DEFAULT choice: {0}'.format(self.col_rad_box_choice))

        match_select_text = 'Choose ONE type of text match\nIf you\'re not sure,'\
        ' choose Regular.'
        self.column_instructions = wx.StaticText(self.panel, label=match_select_text, pos=(200, 290))

        match_list = ['Regular Match', 'Fuzzy Match']
        self.rad_box_match_type = wx.RadioBox(self.panel, label='Type of Text match', pos=(200, 325),
                                              choices=match_list, majorDimension=0, style=wx.RA_SPECIFY_ROWS)

        self.rad_box_match_type.Bind(wx.EVT_RADIOBOX, self.on_radio_box_match_type)
        self.rad_box_match_type_choice = self.rad_box_match_type.GetStringSelection()
        print('Text Match Type DEFAULT admin choice: {0}'.format(self.rad_box_match_type))

        # wx.ID_OK is 5100, wx.OK is 4.
        self.ok_btn = wx.Button(self.panel, id=wx.ID_OK, label='OK', pos=(200, 440))
        self.ok_btn.Bind(wx.EVT_BUTTON, self.on_ok)

        self.cancel_btn = wx.Button(self.panel, id=wx.ID_CANCEL, label='Cancel', pos=(400, 440))
        self.cancel_btn.Bind(wx.EVT_BUTTON, self.on_cancel)

    def on_col_radio_box(self, e):
        self.col_rad_box_choice = e.GetEventObject().GetStringSelection()
        print(self.col_rad_box.GetStringSelection(),' was clicked from Column Priority Radio Box')

    def on_radio_box_match_type(self, e):
        self.rad_box_match_type_choice = e.GetEventObject().GetStringSelection()
        print(self.rad_box_match_type.GetStringSelection(),' was clicked from Text Match Type Radio Box')

    def on_ok(self, event):
        print(event.GetId())
        # Assign attribute so we know OK was pressed
        if self.rad_box_match_type_choice == 'Fuzzy Match':
            self.fuzzy_match = 1

        self.radio_box_pressed_ok_btn = 1
        self.Destroy()

    def on_cancel(self, event):
        print('Cancel button pressed. Closing the Admin choice radio box dialog.')
        self.Destroy()

    def on_radio_group(self, e):
        btn = e.GetEventObject()
        self.radbox_admin_choice = btn.GetStringSelection()
        print(btn.GetStringSelection(), ' was clicked from Radio Group')

    @property
    def admin_choice(self):
        return self.radbox_admin_choice


class EPSGDialog(wx.Dialog): #was wx.Dialog

    # Some borrowed from https://stackoverflow.com/questions/3551249/how-to-make-wx-textentrydialog-larger-and-resizable
    def __init__(self, parent, title, caption):
        style = wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
        super(EPSGDialog, self).__init__(parent, -1, title, style=style)
        text = wx.StaticText(self, -1, caption)
        input = wx.TextCtrl(self, -1, style=wx.TE_LEFT )
        input.SetInitialSize((50, 30))
        buttons = self.CreateButtonSizer(wx.OK | wx.CANCEL )
        # .SetYesNoCancelLabels(wx.ID_OK, 'Skip Shapefile Creation', wx.ID_CANCEL)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(text, 0, wx.ALL, 5)
        sizer.Add(input, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(buttons, 0, wx.EXPAND | wx.ALL, 5)
        self.SetSizerAndFit(sizer)
        self.input = input

    def GetValue(self):
        return self.input.GetValue()


class FuzzyDialog(wx.Dialog):

    def __init__(self, parent, title, caption):
        style = wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER
        super(FuzzyDialog, self).__init__(parent, -1, title, style=style)
        text = wx.StaticText(self, -1, caption)
        input = wx.TextCtrl(self, -1, style=wx.TE_LEFT )
        input.SetInitialSize((40, 30))
        buttons = self.CreateButtonSizer(wx.OK | wx.CANCEL)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(text, 0, wx.ALL, 5)
        sizer.Add(input, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(buttons, 0, wx.EXPAND | wx.ALL, 5)
        self.SetSizerAndFit(sizer)
        self.input = input

    def GetValue(self):
        return self.input.GetValue()


# Works Correctly to show spreadsheet file November 26 2021
class PreviewTable(wx.Panel):

    def __init__(self, parent, df):
        """constructor"""
        wx.Panel.__init__(self, parent=parent)

        self.list_ctrl = wx.ListCtrl(self, size=wx.DefaultSize, #size=(-1, 100),
                                     style=wx.LC_REPORT
                                     )
        self.current_selection = None

        # Create columns first
        col_nums = []
        for idx, col in enumerate(df.columns):
            col_nums.append(idx)
            self.list_ctrl.InsertColumn(idx, col)

        row_num = len(df.index)
        row_count = 0
        df = df.sort_index()

        # Populate rows
        if row_num <= 100:
            for idx, row in df.iterrows():
                self.list_ctrl.InsertItem(idx, str(row[col]))

                # call .InsertItem() for the first column and SetItem()
                # for all the subsequent columns https://realpython.com/python-gui-with-wxpython/
                for col in col_nums:
                    self.list_ctrl.SetItem(idx, col, str(row[col]))
        elif row_num > 100:
            for idx, row in df.iterrows(): # df[df.columns[~df.isnull().all()]].iterrows():
                # Stop after we've shown 100th row in the preview, otherwise it can take long time to load all rows.
                if row_count > 99:
                    break
                self.list_ctrl.InsertItem(idx, str(row[col]))
                for col in col_nums:
                    self.list_ctrl.SetItem(idx, col, str(row[col]) )
                row_count += 1

        sizer = wx.BoxSizer(wx.VERTICAL )
        sizer.Add(self.list_ctrl, 0, wx.ALL | wx.EXPAND, 12)
        self.SetSizer(sizer)


class RedirectText(object):
    # Redirect text to console_text text control element
    def __init__(self, aWxTextCtrl):
        self.out = aWxTextCtrl

    def write(self, string):
        self.out.WriteText(string)


if __name__ == '__main__':
    app = wx.App()
    frame = Frame()
    app.MainLoop()
