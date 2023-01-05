import dash_bootstrap_components as dbc
# Run this app with `python app.py` and
# visit http://127.0.0.1:8050/ in your web browser.
from pathlib import Path
from jupyter_dash import JupyterDash
import dash_uploader as du
import flask
from bs4 import BeautifulSoup
import math
import csv
from dash import Dash, html, dcc, Input, Output, State, callback_context
import plotly.express as px
import pandas as pd
import numpy as np
import os
import openpyxl
import glob
import webbrowser
import signal
from sklearn.metrics import euclidean_distances
import dash_daq as daq
from pathlib import Path
import xlwings as xw
#import pythoncom
import simplekml
from concurrent.futures import ThreadPoolExecutor
from zipfile import ZipFile
import time


sorted_coordinates = []
unsorted_coordinates = []
graph_name = ""
# Set colours so that first plot is shown as unsorted, in orange.
initialFig = px.scatter_geo(color_discrete_sequence=["#ef553c", "#636efa"])
initialFig.update_layout(clickmode='event+select', title=graph_name, title_x=0.55, title_font_family='sans-serif')
suggestedCoordinates = ''
latSuggestion = ''
longSuggestion = ''

clipboardStyle = {"fontSize": 0, 'display': 'inline-block'}
styles = {
    'pre': {
        'border': 'thin lightgrey solid',
        'overflowX': 'scroll'}
}
if os.path.exists('sorted.csv') and os.path.exists('unsorted.csv'):
    os.remove("sorted.csv")
    os.remove("unsorted.csv")

# Initialise Server
server = flask.Flask(__name__)
# Build Components
app = JupyterDash(__name__, title="Coordinate Sorter", server=server, suppress_callback_exceptions=True, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Get current operating directory
dir_path = os.getcwd()
# Create folder path for upload folder
UPLOAD_FOLDER = os.path.join(dir_path, "KML Files")

# Create upload folder for KML files to be stored
try:
    os.mkdir(UPLOAD_FOLDER)
except FileExistsError:
    pass


def get_kml_file():
    """
    This function retrieves the KML files inside the upload folder.
    Returns: A list of kml filenames
    """
    # Find all files ending in .kml, saving their path
    kml_file = glob.glob(UPLOAD_FOLDER + '/*.kml')
    kml_files = []
    # Save filename of the kml files found, into kml_files list
    for file in kml_file:
        # Append just the filename and not path
        kml_files.append(os.path.basename(file))
    return kml_files


# Configure Dash upload component
du.configure_upload(app, UPLOAD_FOLDER, use_upload_id=False)
def RPL_layout():
    # Dash Layout Customisation
    layout = html.Div(children=[
        html.Div(
            # Page title
            [html.H2(children='Coordinate Sorter', style={'display': 'flex'}),

             # Shutdown Server button
             html.Div(html.Button('Shutdown Server', id='Quit_btn', n_clicks=0,

                                  style={'position': 'absolute',
                                         'right': '0',
                                         'bottom': '25%',
                                         'top': '25%',
                                         'background-color': 'red'})),
             html.Div(id='Output_div')
             ], style={
                'display': 'flex',
                'flex-direction': 'row',
                'justify-content': 'center',
                'align-items': 'center',
                'align-content': 'center',
                'position': 'relative',

            }),

        html.Div([

            # Update File List Button
            html.Div([
                html.Button('Update File List', id='update_dropdown_button', n_clicks=0)],
                style={
                    'margin': '10px'
                }),

            # File List Dropdown
            html.Div([
                dcc.Dropdown(get_kml_file(), id='dropdown', placeholder='Select KML File', maxHeight=200, optionHeight=50)],
                style={
                    'width': '15em',
                    'margin': '10px',
                    'text-align': 'left',
                }),

            # Add to map Button
            html.Div([
                html.Button(
                    'Add to Map',
                    id='kml-csv-btn-click',
                    n_clicks=0
                ),

            ]),

        ],
            style={
                'display': 'flex',
                'align-items': 'center',
                'justify-content': 'center'
            }),

        # Upload Files Box
        html.Div([
            du.Upload(
                text='Drag and Drop files here',
                text_completed='Completed: ',
                pause_button=False,
                cancel_button=True,
                max_file_size=1800,  # 1800 Mb
                filetypes=['kml', 'kmz'],
                id='upload-files-div', default_style={'minHeight': 2, 'lineHeight': 2},
            ),
        ],
            style={
                'textAlign': 'center',
                'width': '30%',
                'display': 'inline-block',
            },
        ),

        html.Div(className='row', children=[

            # Div to center Latitude, Longitude, Auto Create rpl, Start Sorting Coordinates, Toggle Button
            html.Div([
                html.Div(
                    id='container-button-kml-csv2',
                    children=' ',
                    className='text'
                ),
            ], className='four columns'),

            #  Starting long and lat
            html.Div([
                # Input lat long
                dcc.Input(
                    id="lat",
                    type="text",
                    placeholder="Latitude",
                    debounce=True, style={'margin': '10px'}),
                dcc.Input(
                    id="lon",
                    type="text",
                    placeholder="Longitude",
                    debounce=True, style={'margin': '10px'}
                ),

                html.Div([

                    # Auto Create Button
                    html.Button(
                        'Auto Create RPL',
                        id='auto-try-btn',
                        n_clicks=0, style={'margin': '10px'}),

                    # Start Sorting Coordinates Button
                    html.Button(
                        'Start Sorting Coordinates',
                        id='sort-coordinates-btn',
                        n_clicks=0, style={'margin': '10px'}),

                    # Enable/Disable stopping at overlap toggle Button
                    daq.ToggleSwitch(
                        id='my-toggle-switch',
                        value=False,
                        color="Blue",
                        label={"label": "Disable/Enable stopping at overlap",
                               "style": {'margin': '10px', 'font-weight': 600, 'font-size': '13px', }},
                        labelPosition="right", style={'margin': '10px', 'text-align': 'center'}
                    ),
                    html.Div(id='my-toggle-switch-output'),
                    html.Br()
                ],
                    style={
                        'display': 'flex',
                        'align-items': 'center',
                        'justify-content': 'center'
                    }),

                # Loading animations
                html.Div([
                    # Loading symbol for Update Output function (Everything)
                    dcc.Loading(
                        id="loading-auto-cre",
                        type="default",
                        children=[html.Div(id="loading-auto-create")],
                        style={
                            'display': 'block'
                        }
                    ),
                    # Loading symbol for KMZ extraction wait
                    dcc.Loading(
                        id="loading-1",
                        type="default",
                        children=[html.Div(id="loading-output")],
                        style={
                            'display': 'block'
                        }
                    )
                ], style={
                    'margin': '10px'
                }),
                # Download RPL Button
                html.Div([
                    html.Button("Download RPL", id="btn_RPL"),
                    dcc.Download(id="download_RPL")
                ], style={
                    'display': 'block',
                    'margin-top': '50px'
                }),

                # Select a file to convert div
                html.Div(
                    id='coordinate-results',
                    children=' '),

                # Suggested lat
                html.Div(
                    id='suggested-lat',
                    children='',
                    style={'display': 'inline-block'}
                ),
                # Clipboard lat
                dcc.Clipboard(
                    id="clipboard-lat",
                    target_id="suggested-lat",
                    title="Copy Latitude",
                    style={
                        "fontSize": 0,
                    },
                ),
                # Suggest long
                html.Div(
                    id='suggested-long',
                    children='',
                    style={'display': 'inline-block',
                           'marginLeft': '10px',
                           'marginRight': '10px'
                           },
                    className='text'
                ),
                # Clipboard long
                dcc.Clipboard(
                    id="clipboard-long",
                    target_id="suggested-long",
                    title="Copy Longitude",
                    style={
                        "fontSize": 0,
                    },
                    className='text'
                ),
            ], className='text four columns'),

            # Click for coordinate
            html.Div([
                html.Div(
                    id='click-coordinates-div',
                    children='Click Coordinate To Get Latitude and Longitude',
                    className='text'
                ),
                html.Br(),
                # Display click lat
                html.Div(
                    id='click-data-lat',
                    children='',
                    style={'display': 'inline-block'},
                    className='text'
                ),
                # Click clipboard lat
                dcc.Clipboard(
                    id="click-clipboard-lat",
                    target_id="click-data-lat",
                    title="copy",
                    style={
                        "fontSize": 0,
                    }
                ),
                html.Br(),
                # Display click long
                html.Div(
                    id='click-data-long',
                    children='',
                    className='text',
                ),
                # Click clipboard long
                dcc.Clipboard(
                    id="click-clipboard-long",
                    target_id="click-data-long",
                    title="copy",
                    style={
                        "fontSize": 0,
                    },
                ),
            ], className='text four columns'),
        ]),
        dcc.Graph(
            figure=initialFig,
            id='Graph',
            style={'display': 'inline-block', 'width': '100%', 'height': '50em'},
        ),
    ], style={'text-align': 'center'})
    return layout

@app.callback(
    Output("Output_div", "children"),
    Input("Quit_btn", "n_clicks")
)
def quit1(n_clicks):
    if n_clicks > 0:
        # Kill program if Quit button pressed
        os.kill(os.getpid(), signal.SIGTERM)
    return

# TODO Add description
@app.callback(
    Output("download_RPL", "data"),
    Input("btn_RPL", "n_clicks"),
    prevent_initial_call=True,
)

def download_RPL(n_clicks):
    global sorted_coordinates
    # Clear coordinates array to avoid added RPL to end of Old RPL if server not restarted
    sorted_coordinates.clear()

    return dcc.send_file(
        'RPL_' + graph_name + '.xls'
    )


# Update Dropdown
def update_options(existing_options):
    """
    This function updates the dropdown box.
    Returns: KML filename
    Params: existing_options: Current list of kml filenames in the dropdown
    """
    # List of kml filenames
    option_name = get_kml_file()
    # Iterate through filenames
    for filename in option_name:
        # Check if filename is in existing_options
        if filename not in existing_options:
            # If filename not in existing_options, append to existing_options
            existing_options.append(filename)
        else:
            # Else just return kml filename
            return option_name

    return option_name


def extract_KML(folder):
    """
    This function extracts kml files from a kmz.
    Returns: Nothing. Saves to upload folder
    Params: folder:Contains the entire contents of a single folder from the kmz
    """
    # Create KML object
    kml = simplekml.Kml()
    # Find name of folder
    folder_name = folder.find('name')
    # Find all coordinates in that folder
    coordinates = folder.find_all('coordinates')

    # For each set of coordinates in that folder, add a new linestring to the kml
    for coor in coordinates:
        kml.newlinestring(name=folder_name, coords=process_coordinate_string(coor.string, True))
    # Remove slashes from folder name to avoid error when creating a path and saving file
    folder_name.string = folder_name.string.replace('/', '')
    folder_name.string = folder_name.string.replace('\\', '')

    # Set KML name and save directory
    kmlname = os.path.join(UPLOAD_FOLDER, (folder_name.string + ".kml"))
    # Save KML
    kml.save(kmlname)


@app.callback(Output('loading-output', 'children'),
              [Input('upload-files-div', 'isCompleted')],
              [State('upload-files-div', 'fileNames')],)
def upload_files(isCompleted, fileNames):
    """
    This function uploads and detects kmz files.
    Params: isCompleted: True or False from Dash Upload component.
            fileNames: List of filenames uploaded from Dash Upload component
    Returns: Empty String. Responsible for starting extraction of KMZ file.

    """
    if not isCompleted:
        # Returning list of length 2 to avoid loading error
        return ["", ""]
    if fileNames is not None:
        out = []
        # For each filename in fileNames check if .kmz.
        for filename in fileNames:
            if os.path.splitext(filename)[1] == ".kmz":
                # Unzip kmz to reveal doc.kml
                kmz = ZipFile(os.path.join(UPLOAD_FOLDER, filename), 'r')
                # Open doc.kml
                kml_filename = kmz.open('doc.kml', 'r')
                # Parse kml file
                s = BeautifulSoup(kml_filename, 'xml')
                # Find all folders
                folder_list = s.find_all('Folder')
                # Execute extract_KML() function for each folder in a different thread, parallel tasking, improving performance.
                with ThreadPoolExecutor() as executor:
                    [executor.submit(extract_KML(folder=folder)) for folder in folder_list]
            else:
                # Save file
                file = Path(UPLOAD_FOLDER) / filename
                out.append(file)
        return ["", ""]
    return [html.Ul(html.Li("No Files Uploaded Yet!"))]


@app.callback(
    Output('click-data-long', 'children'),
    Output('click-data-lat', 'children'),
    Output('click-clipboard-long', 'style'),
    Output('click-clipboard-lat', 'style'),
    Input('Graph', 'clickData'))
def display_click_data(clickData):
    """
    This function returns the longitude and latitude of the point the user clicks on the map.
    Params: clickData: Longitude and Latitude of point clicked.
    Returns: Longitude and Latitude of the user selected point.
    """
    if clickData is None:
        return ' ', '', {"fontSize": 0, 'display': 'inline-block'}, {"fontSize": 0, 'display': 'inline-block'}
    else:
        return (
            clickData["points"][0]["lon"], clickData["points"][0]["lat"], {"fontSize": 20, 'display': 'inline-block'},
            {"fontSize": 20, 'display': 'inline-block'})


@app.callback(
    Output('my-toggle-switch-output', 'children'),
    Output('btn_RPL', 'style'),
    Output('dropdown', 'options'),
    Output('coordinate-results', 'children'),
    Output('clipboard-lat', 'style'),
    Output('suggested-lat', 'children'),
    Output('clipboard-long', 'style'),
    Output('suggested-long', 'children'),
    Output('Graph', 'figure'),
    Output('sort-coordinates-btn', 'disabled'),
    Output('auto-try-btn', 'disabled'),
    Output('loading-auto-create', 'children'),
    Input('dropdown', 'value'),
    Input('kml-csv-btn-click', 'n_clicks'),
    Input('sort-coordinates-btn', 'n_clicks'),
    Input("lat", "value"),
    Input("lon", "value"),
    [State('dropdown', 'options')],
    Input('update_dropdown_button', 'n_clicks'),
    Input('auto-try-btn', 'n_clicks'),
    Input('my-toggle-switch', 'value')

)
def update_output(value, KML_clicks, sort_clicks, lat, lon, existing_options, dropdown_clicks, auto_try_clicks, toggle):
    """
    This function controls almost every output onto the Dash display. Any click the user makes, flows through this function.
    Params: value: Selected kml from dropdown. Contains filename of kml. If not selected = None.
            KML_clicks: Number of times the Add to Map button has been pressed.
            sort_clicks: Number of times the Start Sorting Coordinates button has been preseed.
            lat: Contains the current Latitude in the latitude input box.
            lon: Contains the current Longitude in the Longitude input box.
            existing_options: Contains the list of options currently in the dropdown.
            dropdown_clicks: Number of times the Update File List button has been pressed.
            auto_try_clicks: Number of times Auto Create RPL button has been pressed.
            toggle: True or False for the Disable/Enable stopping at overlap toggle. Default = False.
    Returns: All of the above parameters with the corresponding value depending on the action that called the function.
    """
    global initialFig
    global graph_name
    global suggestedCoordinates
    global latSuggestion
    global longSuggestion
    global clipboardStyle
    global sorted_coordinates
    global unsorted_coordinates

    # Find which action has been changed/pressed.
    changed_id = [p['prop_id'] for p in callback_context.triggered][0]

    # Initialise the current kml files.
    kmlfiles = update_options(existing_options)

    # Check if Add to Map button has been pressed or Disable/Enable stopping at overlap has changed
    if 'kml-csv-btn-click' in changed_id or 'my-toggle-switch' in changed_id:
        if value is None:
            # Send "Select a file to convert" message to the user
            return toggle, {
                'display': 'none'}, kmlfiles, 'Select a file to convert', clipboardStyle, longSuggestion, clipboardStyle, latSuggestion, initialFig, False, False, ""
        # Send selected kml filename to convertKMLToCSV function. Creates a csv containing all the coordinates of kml file
        convertKMLToCSV(value)
        # Set column names for dataframe
        colnames = ['latitude', 'longitude']
        # Save coordinates into a dataframe using the csv convertKMLToCSV function created
        df = pd.read_csv('coordinates.csv', names=colnames, header=None)
        # Initialise map with the dataframe of coordinates
        fig = px.scatter_geo(df, lat='latitude', lon='longitude', color_discrete_sequence=["#ef553c", "#636efa"])
        # Set graph characteristics, like title.
        fig.update_layout(title=graph_name, title_x=0.55, title_font_family='sans-serif')
        # Set size of marker
        fig.update_traces(marker=dict(size=4))
        # Center map view to the cable that has been plotted
        fig.update_geos(fitbounds="locations")
        # Apply fig to the global figure
        initialFig = fig
        return toggle, {
            'display': 'none'}, kmlfiles, '', clipboardStyle, longSuggestion, clipboardStyle, latSuggestion, fig, False, False, ""

    # Check if Sort Coordinates button or Auto Create RPL button has been pressed
    if 'sort-coordinates-btn' in changed_id or 'auto-try-btn' in changed_id:
        # Check if unsorted_coordinates is empty
        if not unsorted_coordinates:
            # Send, "Add to Map First", to the user as the coordinates have not been sorted yet
            return toggle, {
                'display': 'none'}, kmlfiles, 'Add to Map First', clipboardStyle, longSuggestion, clipboardStyle, latSuggestion, px.scatter_geo(), False, False, ""
        # Check if Lon and Lat are not None
        if lon and lat:
            # Check lon and lat are numbers
            try:
                int(lon)
                int(lat)
            except ValueError:
                try:
                    float(lon)
                    float(lat)
                except ValueError:
                    return toggle, {
                        'display': 'none'}, kmlfiles, 'Enter Int or Float', clipboardStyle, latSuggestion, clipboardStyle, longSuggestion, initialFig, False, False, ""
            # Save order_section result
            result = order_section(lat, lon)
            # Save coordinates to be used in next loop
            prev_coor = [(lon, lat)]
            # Save result to be used in next loop
            prev_result = result
            l = 0
            # TODO why does it return a list when there is an overlap
            # If there is an overlap, result returns a list
            while type(result) == list:
                # Create unsorted dataframe from unsorted.csv
                unsorteddf = pd.read_csv(f'unsorted.csv')
                # Create sorted dataframe from sorted.csv
                sorteddf = pd.read_csv(f'sorted.csv')
                # Combine sorted and unsorted dataframes
                df_res = pd.concat([sorteddf, unsorteddf])
                # Plot both sorted and unsorted dataframes
                fig = px.scatter_geo(df_res, lat='Latitude', lon='Longitude', hover_name="Colour", color="Colour",
                                     color_discrete_sequence=["#636efa", "#ef553c"], labels={"Colour": "Key"})
                fig.update_layout(title=graph_name, title_x=0.55, title_font_family='sans-serif')
                fig.update_geos(fitbounds="locations")
                fig.update_traces(marker=dict(size=4))
                initialFig = fig
                # Get suggested next coordinates by sending lat and lon
                suggestedNextCoordinates = findNextPoint(result[0], result[1])
                # The style of the coordinates to be copied by the user
                clipboardStyle = {"fontSize": 20, 'display': 'inline-block'}
                # Save suggested next coordinates individually as lat and long
                latSuggestion = suggestedNextCoordinates[0]
                longSuggestion = suggestedNextCoordinates[1]
                # Create response message for suggested coordinates to be output to the user.
                suggestedCoordinates = (
                    'Found gap, last known coordinates: {} , {} '.format(result[0], result[1]), html.Br(),
                    ' Suggest coordinates:')
                # Add coordinates to previous coordinates list
                prev_coor.append((result[0], result[1]))
                # Get new result based on the suggestion of next coordinates
                result = order_section(latSuggestion, longSuggestion)

                # Cycle through previous coordinates until no longer an overlap if auto try button pressed
                if prev_result == result:
                    if 'auto-try-btn' in changed_id:

                        # Decrease counter to count back through previous coordinates
                        l -= 1
                        # Attempt to order section with previous coordinates, until no overlap.
                        result = order_section(prev_coor[l][0], prev_coor[l][1])
                        # Prevent Infinite Loop
                        if prev_coor[l] == prev_coor[l - 1]:
                            # Delete previous entry as duplicate
                            del prev_coor[-1]
                            return toggle, {
                                'display': 'none'}, kmlfiles, suggestedCoordinates, clipboardStyle, latSuggestion, clipboardStyle, longSuggestion, fig, False, False, ""

                if toggle == True:
                    # Indicates an overlap
                    if prev_result == result:
                        # Stop at the overlap and ask user to enter next coordinates
                        return toggle, {
                            'display': 'none'}, kmlfiles, suggestedCoordinates, clipboardStyle, latSuggestion, clipboardStyle, longSuggestion, fig, False, False, ""
                    # Update previous result
                    prev_result = result
                    # Increment counter
                    l += 1
                elif toggle == False and 'auto-try-btn' not in changed_id:
                    # Update previous result
                    prev_result = result
                    # Increment counter
                    l += 1
                    return toggle, {
                        'display': 'none'}, kmlfiles, suggestedCoordinates, clipboardStyle, latSuggestion, clipboardStyle, longSuggestion, fig, False, False, ""
                else:
                    # Update previous result
                    prev_result = result
                    # Increment counter
                    l += 1
            else:
                # Save sorted.csv into a dataframe
                sorteddf = pd.read_csv(f'sorted.csv')
                # Initialise graph with sorted.csv
                fig = px.scatter_geo(sorteddf, lat='Latitude', lon='Longitude', hover_name="Colour", color="Colour",
                                     color_discrete_sequence=["#636efa", "#ef553c"], labels={"Colour": "Key"})
                fig.update_layout(title=graph_name, title_x=0.55)
                fig.update_traces(marker=dict(size=4))
                fig.update_geos(fitbounds="locations")
                # Save sorted.csv into a dataframe without colour column
                sortedCoordinates = (pd.read_csv(f'sorted.csv')).drop(['Colour'], axis=1)
                # Create RPL with sorted coordinates
                creatRPLTemplate(sortedCoordinates)
                # Convert dataframe into numpy array
                sortedCoordinates = sortedCoordinates.to_numpy()
                # Create a csv holding the sorted coordinates
                with open("sortedCoordinates.csv", "w") as my_csv:
                    csvWriter = csv.writer(my_csv, delimiter=',')
                    for x in sortedCoordinates:
                        csvWriter.writerow(x)
                # Delete sorted.csv
                os.remove("sorted.csv")
                os.remove("sortedCoordinates.csv")
                return toggle, {'display': 'inline-block', }, kmlfiles, result, {"fontSize": 0,
                                                                                 'display': 'inline-block'}, '', {
                           "fontSize": 0, 'display': 'inline-block'}, '', fig, True, True, ""
        else:
            return toggle, {
                'display': 'none'}, kmlfiles, 'Please enter Coordinates', clipboardStyle, longSuggestion, clipboardStyle, latSuggestion, initialFig, False, False, ""
    else:
        return toggle, {
            'display': 'none'}, kmlfiles, suggestedCoordinates, clipboardStyle, latSuggestion, clipboardStyle, longSuggestion, initialFig, False, False, ""


def creatRPLTemplate(coordinates):
    """
    This function creates the final RPL files to be downloaded by the user.
    Params: coordinates: Dataframe of coordinates
    Returns: Nothing. Directly creates RPL as .xls in the directory.
    """
    global graph_name
    global dir_path
    # Load RPL template
    wrkbk = openpyxl.load_workbook('RPLTemplate.xlsx')
    # Load current sheet in RPLTemplate, there is only one sheet
    sh = wrkbk.active
    # Template data starts from row 7 onwards as the template contains a header
    i = 7
    # Loop through dataframe of coordinates
    for index, row in coordinates.iterrows():
        # Split integer and fractional parts of longitude and latitude
        latFrac, LatWhole = math.modf(row['Latitude'])
        longFrac, longWhole = math.modf(row['Longitude'])
        # lat
        sh.cell(row=i, column=3).value = LatWhole
        sh.cell(row=i, column=4).value = latFrac * 60
        sh.cell(row=i, column=5).value = 'N' if LatWhole > 0 else 'S'
        # long
        sh.cell(row=i, column=6).value = longWhole
        sh.cell(row=i, column=7).value = longFrac * 60
        sh.cell(row=i, column=8).value = 'E' if longWhole > 0 else 'W'
        # Increment row by 2 as next row in RPLTemplate should not be filled
        i += 2
    # Calculate number of rows left
    num_rows_left = sh.max_row - i
    # Delete rows left
    sh.delete_rows(i, (num_rows_left + 1))

    # ------------ Rename sheet name with segment name --------------

    # Find segment name
    # TODO implement more ways of finding segment name from KML.
    if "seg" in graph_name or "Seg" in graph_name:
        segment_name = graph_name.split("seg", 1)[1] or graph_name.split("Seg", 1)[1]
    else:
        segment_name = "Unknown Segment"

    # Change worksheet title
    sh.title = segment_name

    # ------------ Change references from old worksheet name to new, fixing broken links --------------

    # Iterate through all name ranges
    for i in range(len(wrkbk.defined_names.definedName)):
        # Try to access destinations where a worksheet will be referenced, if not pass
        try:
            # If destination of define ranges equals old worksheet name
            if list(wrkbk.defined_names.definedName[i].destinations)[0][0] == 'Segment_Name':
                # Set old name range name
                old_range_name = str(list(wrkbk.defined_names.definedName[i])[0][1])
                # Set old name range attr text
                old_attr_text = list(wrkbk.defined_names.definedName[i].destinations)[0][1]
                # Delete old name range before we create a new one
                del wrkbk.defined_names.definedName[i]
                # Create new attr text for the new name range
                new_attr_text = "'" + segment_name + "'" + "!" + old_attr_text
                # Create the new name range object with old name range name as the name and new attr text
                A_New_Range = openpyxl.workbook.defined_name.DefinedName(old_range_name, attr_text=new_attr_text)
                # Add new name range to workbook with the updated reference to the new sheet name.
                wrkbk.defined_names.append(A_New_Range)

        except IndexError:
            # If destinations were none, pass.
            continue

    # Save workbook
    wrkbk.save('RPL_' + graph_name + '.xlsx')

    # Get filepath of new RPL
    filepath = os.path.join(dir_path, ('RPL_' + graph_name + '.xlsx'))
    # Convert to .xls file
    pre, ext = os.path.splitext(filepath)
    try:
        os.rename(filepath, pre + '.xls')
    except FileExistsError:
        # Delete file if already exists
        os.remove(os.path.join(dir_path, ('RPL_' + graph_name + '.xls')))
        os.rename(filepath, pre + '.xls')


def process_coordinate_string(coor_str, kmz):
    """
    This function breaks the string of coordinates into [Lat,Lon,Lat,Lon...] for a CSV row
    Params: coor_str: String of coordinates
            kmz: If KMZ or not. bool, True or False
    Returns: A list of lists containing latitude and longitude coordinates. [[lat, lon],[lat, lon]]
    """

    # Clean string of \n and \t
    coor_str = coor_str.strip()

    # TODO IF ERROR, ADD ERROR IN FOR USER WITHOUT HAVING CALLBACK ON.
    unclean_space_splits = coor_str.split(" ")
    space_splits = []
    for x in unclean_space_splits:
        # print("x= ", x)
        # for char in range(len(x)):
        # print(x[char])
        # print("x[1:2]= ", x[1:2])
        # print("x[0:1]= ", x[0:1])

        if x[1:2].isdigit() or x[0:1].isdigit():
            space_splits.append(x)
    # space_splits = [x for x in unclean_space_splits if x[1:2].isdigit()]

    # print(unclean_space_splits[3][:5])
    # print(type(unclean_space_splits[0]))
    rows = len(space_splits)
    ret = [[0 for x in range(2)] for y in range(rows)]
    for count, value in enumerate(space_splits, start=0):
        comma_split = value.split(',')
        if kmz is True:
            ret[count][1] = (comma_split[1])  # lat
            ret[count][0] = (comma_split[0])  # lng
        else:
            ret[count][0] = (comma_split[1])  # lat
            ret[count][1] = (comma_split[0])  # lng
    return ret


# CALLBACK error value
def convertKMLToCSV(value):
    global graph_name
    global sorted_coordinates
    global unsorted_coordinates
    unsorted_coordinates.clear()
    if os.path.exists("coordinates.csv"):
        os.remove("coordinates.csv")
    kml_filename = os.path.join(UPLOAD_FOLDER, value)

    with open(kml_filename, 'r') as f:
        s = BeautifulSoup(f, 'xml')
        # If name is none, name graph filename
        if s.find('Schema'):
            graph_name = s.find('Schema')['name']
        else:
            # Save graph name as file name.
            # Save first index 0 from split as index 1 is the file extension, .kml
            graph_name = os.path.basename(os.path.splitext(kml_filename)[0])
        for coords in s.find_all('coordinates'):
            file = open(r'coordinates.csv', 'a', newline='')
            with file:
                write = csv.writer(file)
                write.writerows([" "])
                section = process_coordinate_string(coords.string, False)
                write.writerows(section)
                unsorted_coordinates.append(section)


def appendColumnToArray(coordinates, column):
    full = np.full((len(coordinates), 1), column)
    arr = np.array(coordinates)
    return np.hstack((arr, full))


def order_section(lat, lon):
    global sorted_coordinates
    global unsorted_coordinates
    starting = [lat, lon]
    sorted = coordinates_found = coordinates_added = False
    section_start = len(sorted_coordinates)
    starting_coordinates = starting.copy()
    while sorted == False:
        coordinates_found = False
        for section in unsorted_coordinates:
            if section[0][0] == starting_coordinates[0] and section[0][1] == starting_coordinates[1]:
                starting_coordinates = (section[(len(section) - 1)]).copy()
                sorted_coordinates.append(section[1:])
                coordinates_found, coordinates_added = True, True
                unsorted_coordinates.remove(section)

            elif section[(len(section) - 1)][0] == starting_coordinates[0] and section[(len(section) - 1)][1] == \
                    starting_coordinates[1]:
                section.reverse()
                starting_coordinates = (section[(len(section) - 1)]).copy()
                sorted_coordinates.append(section[1:])
                coordinates_found, coordinates_added = True, True
                unsorted_coordinates.remove(section)

        if coordinates_found == False:
            if coordinates_added == True:
                sorted_coordinates[section_start].insert(0, starting)
                section_start = len(sorted_coordinates)
            headings = ["Latitude", "Longitude", "Colour"]
            # unsorted and sorted to csv
            with open("sorted.csv", "w+") as my_csv:
                csvWriter = csv.writer(my_csv, delimiter=',')
                csvWriter.writerow(headings)
                for section in sorted_coordinates:
                    csvWriter.writerows(appendColumnToArray(section, "Sorted"))
                    # add to dataframe unsorted
            with open("unsorted.csv", "w+") as my_csv:
                csvWriter = csv.writer(my_csv, delimiter=',')
                csvWriter.writerow(headings)
                for section in unsorted_coordinates:
                    csvWriter.writerows(appendColumnToArray(section, "Unsorted"))
                    # add to df sorted
            # use sorted and unsorted csv instead
            if len(unsorted_coordinates) == 0:
                sorted = True
                if os.path.exists("unsorted.csv"):
                    os.remove("unsorted.csv")
                return "Coordinates sorted, ", "RPL_" + graph_name + ".xls created"
            else:
                return [starting_coordinates[0], starting_coordinates[1]]


def findNextPoint(lat, lon):
    x = np.array([[lat, lon]])
    route = pd.read_csv(f'unsorted.csv', dtype=str)
    route = route.drop(['Colour'], axis=1)
    route = route.to_numpy()
    dist = euclidean_distances(x, route)
    index = np.argmin(euclidean_distances(x, route))
    return route[index]

