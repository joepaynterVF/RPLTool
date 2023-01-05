import base64
import os
import zipfile
from os.path import basename
from urllib.parse import quote as urlquote
from bs4 import BeautifulSoup
from flask import Flask, send_from_directory
from dash import dcc, html, Input, Output
import simplekml
from concurrent.futures import ThreadPoolExecutor
from zipfile import ZipFile
from RPLTool import app, server, process_coordinate_string, quit1
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
                    id="loading-11",
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
    kmlname = os.path.join(UPLOAD_DIRECTORY, (folder_name.string + ".kml"))
    # Save KML
    kml.save(kmlname)


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
