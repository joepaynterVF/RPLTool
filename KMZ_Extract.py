import base64
import os
import zipfile
from os.path import basename
from urllib.parse import quote as urlquote
from bs4 import BeautifulSoup
from flask import send_from_directory
from dash import dcc, html, Input, Output
import simplekml
from concurrent.futures import ThreadPoolExecutor
from zipfile import ZipFile
from RPLTool import app, server
import dash_bootstrap_components as dbc

# Get current operating directory
dir_path = os.getcwd()
# Create folder path for upload folder
UPLOAD_DIRECTORY = os.path.join(dir_path, "KMZ Extract Files")

if not os.path.exists(UPLOAD_DIRECTORY):
    os.makedirs(UPLOAD_DIRECTORY)


@server.route("/download/<path:path>")
def download(path):
    """Serve a file from the upload directory."""
    return send_from_directory(UPLOAD_DIRECTORY, path, as_attachment=True)


def create_page_KMZ_extract():
    """
    This function returns the layout for the KMZ_Extract page used by Dash.
    :return: Dash Layout
    """
    layout = html.Div(
        [
            html.Div(
                # Page title
                [html.H2(children='KMZ Extractor', style={'display': 'flex'}),

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

            # Upload Box
            dcc.Upload(
                id="upload-data",
                children=html.Div(
                    ["Drag and Drop a .kmz file here"]
                ),
                style={
                    "width": "30%",
                    "height": "60px",
                    "lineHeight": "60px",
                    "borderWidth": "1px",
                    "borderStyle": "dashed",
                    "borderRadius": "10px",
                    "textAlign": "center",
                    "margin": "10px",
                    'display': 'inline-block',
                },
                multiple=True,
            ),
            html.Div([
                # Loading symbol for KMZ extraction wait
                dcc.Loading(
                    id="loading-11",
                    type="default",
                    children=[html.Div(id="loading-output-KMZ")],
                    style={
                        'display': 'inline-block'
                    }),
                dcc.Loading(
                    id="loading-12",
                    type="default",
                    children=[html.Div(id="loading-output-download-all")],
                    style={
                        'display': 'inline-block'
                    })
            ],
                style={
                    'margin-top': '20px'
                }
            ),

            html.Div(
                [
                    # Search box for file list
                    dbc.Input(id="input", placeholder="Search File List", type="text",
                              style={"margin": "auto", "width": "15%", "text-align": "center"}),
                ], style={
                    "margin-top": "50px"
                }
            ),
            html.H6("File List"),
            html.Div(
                # Download All Button
                ['Click a file to download or   ',
                 html.Button('Download All (.zip)', id='download_all_btn', n_clicks=0),
                 dcc.Download(id="download_all")],

                className='text',
                style={
                    'margin': '10px'
                }
            ),

            html.Ul(id="file-list"),
        ],
        style={"text-align": "center"},
    )
    return layout


@app.callback(
    Output("download_all", "data"),
    Output("loading-output-download-all", "children"),
    Input('download_all_btn', 'n_clicks'),
    prevent_initial_call=True,
)
def download_all(n_clicks):
    """
    This function compresses all the files into a zip and creates a download link for the user.
    Parameters: n_clicks: Number of times the download button has been pressed.
    Return: KMZ_Extract.zip containing all the files.
    """
    global file_list

    # Create zip file containing all individual kml files
    with ZipFile("KMZ_Extract.zip", 'w') as zipObj:
        # Iterate through each file in the file_list
        for file in file_list:
            # Write each KML file into the .zip folder
            filePath = os.path.join(UPLOAD_DIRECTORY, file)
            zipObj.write(filePath, basename(filePath))
    return dcc.send_file("KMZ_Extract.zip"), ("", "")


def save_file(name, content):
    """
    Decode and store a file uploaded with Plotly Dash.
    """
    data = content.encode("utf8").split(b";base64,")[1]
    with open(os.path.join(UPLOAD_DIRECTORY, name), "wb") as fp:
        fp.write(base64.decodebytes(data))


def uploaded_files():
    """List the files in the upload directory."""
    files = []
    for filename in os.listdir(UPLOAD_DIRECTORY):
        path = os.path.join(UPLOAD_DIRECTORY, filename)
        if os.path.isfile(path):
            files.append(filename)
    return files


def file_download_link(filename):
    """Create a Plotly Dash 'A' element that downloads a file from the app."""
    location = "/download/{}".format(urlquote(filename))
    return html.A(filename, href=location)

def process_coordinate_string(coor_str, kmz):
    """
    This function breaks the string of coordinates into [Lat,Lon,Lat,Lon...] for a CSV row
    Params: coor_str: String of coordinates
            kmz: If KMZ or not. bool, True or False
    Returns: A list of lists containing latitude and longitude coordinates. [[lat, lon],[lat, lon]]
    """

    # Clean string of \n and \t
    coor_str = coor_str.strip()

    # Split string by spaces
    unclean_space_splits = coor_str.split(" ")
    space_splits = []
    # Loop through unclean_space_splits
    for x in unclean_space_splits:
        # Check if a section is a digit
        if x[1:2].isdigit() or x[0:1].isdigit():
            # If digit add to space_splits
            space_splits.append(x)

    # Set number of rows as length of space_splits list
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
    # Initialise Variables
    placemark_coordinates = []
    system_name = []
    ship_operation = []
    operation_date = []
    cable_type = []
    slack_percent = []
    segment_name = []
    system_type = []
    installation_year = []
    out_of_service_year = []
    allsimpledataobjs = {}

    if folder.find_all('Placemark'):
        # Find all placemarks in folder
        placemarks = folder.find_all('Placemark')
        # For each placemark object in each placemark
        for placemarkobj in placemarks:
            if placemarkobj.find_all('SimpleData'):
                # Find all simpledata tags
                simpledata = placemarkobj.find_all("SimpleData")

                # Cycle through each simpledata tag, saving to allsimpledataobjs dict
                for simpledataobj in simpledata:
                    # Save each simpledata attribute name and the associated value
                    allsimpledataobjs[simpledataobj.get('name')] = simpledataobj.text

                # Cycle through each simpledata attribute adding the value to associate attribute array
                if "system_name" in allsimpledataobjs:
                    systemname = allsimpledataobjs["system_name"]
                    # Clean value so XML does not throw an error.
                    systemname = clean(systemname)
                    system_name.append(systemname)
                else:
                    # If attribute not in simpledata array within KML, append empty field
                    system_name.append("")
                if "ship_operation" in allsimpledataobjs:
                    shipoperation = allsimpledataobjs["ship_operation"]
                    shipoperation = clean(shipoperation)
                    ship_operation.append(shipoperation)
                else:
                    ship_operation.append("")
                if "operation_date" in allsimpledataobjs:
                    operationdate = allsimpledataobjs["operation_date"]
                    operationdate = clean(operationdate)
                    operation_date.append(operationdate)
                else:
                    operation_date.append("")
                if "cable_type" in allsimpledataobjs:
                    cable_type.append(allsimpledataobjs["cable_type"])
                else:
                    cable_type.append("")
                if "slack_percent" in allsimpledataobjs:
                    slackpercent = allsimpledataobjs["slack_percent"]
                    slackpercent = clean(slackpercent)
                    slack_percent.append(slackpercent)
                else:
                    slack_percent.append("")
                if "segment_name" in allsimpledataobjs:
                    segmentname = allsimpledataobjs["segment_name"]
                    segmentname = clean(segmentname)
                    segment_name.append(segmentname)
                else:
                    segment_name.append("")
                if "system_type" in allsimpledataobjs:
                    systemtype = allsimpledataobjs["system_type"]
                    systemtype = clean(systemtype)
                    system_type.append(systemtype)
                else:
                    system_type.append("")
                if "installation_year" in allsimpledataobjs:
                    installationyear = allsimpledataobjs["installation_year"]
                    installationyear = clean(installationyear)
                    installation_year.append(installationyear)
                else:
                    installation_year.append("")
                if "out_of_service_year" in allsimpledataobjs:
                    outofserviceyear = allsimpledataobjs["out_of_service_year"]
                    outofserviceyear = clean(outofserviceyear)
                    out_of_service_year.append(outofserviceyear)
                else:
                    out_of_service_year.append("")
                # Save coordinates within this placemark
                placemark_coordinates.append(placemarkobj.find("coordinates"))
                # Clear allsimpledataobjs, ready for next placemark within folder
                allsimpledataobjs.clear()

        i = 0
        while i <= len(placemark_coordinates):

            try:
                # For each set of coordinates in each placemark, add a new linestring to the kml
                linestring = kml.newlinestring(name=folder_name,coords=process_coordinate_string(placemark_coordinates[i].string, True))
                # For each set of simpledata attributes in each placemark, add new simpledata field with corrospoding value
                linestring.extendeddata.schemadata.newsimpledata('system_name', system_name[i])
                linestring.extendeddata.schemadata.newsimpledata('ship_operation', ship_operation[i])
                linestring.extendeddata.schemadata.newsimpledata('operation_date', operation_date[i])
                linestring.extendeddata.schemadata.newsimpledata('cable_type', cable_type[i])
                linestring.extendeddata.schemadata.newsimpledata('slack_percent', slack_percent[i])
                linestring.extendeddata.schemadata.newsimpledata('segment_name', segment_name[i])
                linestring.extendeddata.schemadata.newsimpledata('system_type', system_type[i])
                linestring.extendeddata.schemadata.newsimpledata('installation_year', installation_year[i])
                linestring.extendeddata.schemadata.newsimpledata('out_of_service_year', out_of_service_year[i])
            except IndexError:
                break
            # Increment i to cycle through placemark_coordinates and all the arrays
            i += 1

    # Remove chars from folder name to avoid error when creating a path and saving file
    folder_name.string = folder_name.string.replace('/', '')
    folder_name.string = folder_name.string.replace('\\', '')
    folder_name.string = folder_name.string.replace('&', '')
    folder_name.string = folder_name.string.replace('<', '')
    folder_name.string = folder_name.string.replace('>', '')
    folder_name.string = folder_name.string.replace('"', '')
    folder_name.string = folder_name.string.replace("'", '')

    # Set KML name and save directory
    kmlname = os.path.join(UPLOAD_DIRECTORY, (folder_name.string + ".kml"))
    # Save KML
    kml.save(kmlname)

def clean(var):
    """
    This function cleans the variable of /, \\, &, <, >, ", '
    :param var: Variable being passed in
    :return: var, clean variable being returned.
    """
    var = var.replace('/', '')
    var = var.replace('\\', '')
    var = var.replace('&', '')
    var = var.replace('<', '')
    var = var.replace('>', '')
    var = var.replace('"', '')
    var = var.replace("'", '')
    return var

@app.callback(
    Output('loading-output-KMZ', 'children'),
    Output("file-list", "children"),
    [Input("upload-data", "filename"), Input("upload-data", "contents")], [Input("input", "value")]
)
def update_output(uploaded_filenames, uploaded_file_contents, value):
    """Save uploaded files and regenerate the file list."""

    global file_list

    if uploaded_filenames is not None and uploaded_file_contents is not None:
        for name, data in zip(uploaded_filenames, uploaded_file_contents):
            # Check file type is .kmz
            if ".kmz" in name:
                # Save file
                save_file(name, data)
                # Check for bad filetype again, Catches exmaplefile.kmz.kml
                try:
                    # Unzip kmz to reveal doc.kml
                    kmz = ZipFile(os.path.join(UPLOAD_DIRECTORY, name), 'r')
                except zipfile.BadZipfile:
                    # Remove incorrect file previously saved
                    os.remove(os.path.join(UPLOAD_DIRECTORY, name))
                    return ["Incorrect file type. Only .kmz supported", html.Li("No files yet!")]
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
                return ["Incorrect file type. Only .kmz supported", html.Li("No files yet!")]
    files = uploaded_files()
    file_list = []
    if len(files) == 0:
        return ["", html.Li("No files yet!")]
    else:
        if not value:
            file_list = files
            return ("", ""), [html.Li(file_download_link(filename)) for filename in files]
        else:
            # If search term entered, iterate through files list
            for filename in files:
                # Check if lowercase search term appears in filename
                if value.lower() in filename.lower():
                    # If so, append to file list
                    file_list.append(filename)
                else:
                    pass
            # Return filenames from file_list as a link to download
            return ("", ""), [html.Li(file_download_link(filename)) for filename in file_list]
