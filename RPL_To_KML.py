import base64
from flask import send_from_directory
from RPLTool import app, server
from dash import html, dcc, Input, Output, callback_context
import pandas as pd
import os
import glob
import simplekml

# Get current operating directory
dir_path = os.getcwd()
# Create folder path for upload folder
UPLOAD_DIRECTORY = os.path.join(dir_path, "RPL to KML")

if not os.path.exists(UPLOAD_DIRECTORY):
    os.makedirs(UPLOAD_DIRECTORY)


@server.route("/download/<path:path>")
def downloadRPL(path):
    """Serve a file from the upload directory."""
    return send_from_directory(UPLOAD_DIRECTORY, path, as_attachment=True)


def create_page_RPL_to_KML():
    layout = html.Div([
        html.Div(
            # Page title
            [html.H2(children='RPL to KML', style={'display': 'flex'}),

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
            id="upload-data-RPL",
            children=html.Div(
                ["Drag and Drop a .xls or .xlsx file here"]
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
            multiple=True
        ),
        html.Div(
            id="upload-RPL-output",

        ),
        html.Div([
            # Loading symbol for Upload and RPL to KML button
            dcc.Loading(
                id="loading-12",
                type="default",
                children=[html.Div(id="loading-output-RPL-KML")],
                style={
                    'display': 'inline-block'
                })
        ],
            style={
                'margin-top': '20px'
            }
        ),
        html.Button('Convert RPL to KML', id='RPL_to_KML_btn', n_clicks=0, style={"margin-top": "30px"}),

        html.Div([
            html.Button("Download KML", id="btn_KML"),
            dcc.Download(id="download_KML")
        ], style={
            'display': 'block',
            'margin-top': '50px'
        }),
    ],
        style={"text-align": "center"}, )
    return layout


@app.callback(
    Output("download_KML", "data"),
    Input('btn_KML', 'n_clicks'),
    prevent_initial_call=True,
)
def download_KML(n_clicks):
    global kmlname

    return dcc.send_file(kmlname)


def get_RPL_file():
    """
    This function retrieves the RPL files inside the upload folder.
    Returns: A list of RPL filenames
    """

    # Find all files ending in .xls and .xlsx, saving their path
    xls_file = glob.glob(UPLOAD_DIRECTORY + '/*.xls')
    xlsx_file = glob.glob(UPLOAD_DIRECTORY + '/*.xlsx')
    RPL_files = []
    # Save filename of the .xls files found, into xls_files list
    for file in xls_file:
        # Append just the filename and not path
        RPL_files.append(os.path.basename(file))
    for file in xlsx_file:
        # Append just the filename and not path
        RPL_files.append(os.path.basename(file))

    return RPL_files


def save_file(name, content):
    """Decode and store a file uploaded with Plotly Dash."""
    data = content.encode("utf8").split(b";base64,")[1]
    with open(os.path.join(UPLOAD_DIRECTORY, name), "wb") as fp:
        fp.write(base64.decodebytes(data))


@app.callback(
    Output('btn_KML', 'style'),
    Output('upload-RPL-output', 'children'),
    Output("loading-output-RPL-KML", "children"),
    [Input("upload-data-RPL", "filename"), Input("upload-data-RPL", "contents")], Input("RPL_to_KML_btn", "n_clicks")
)
def update_output(uploaded_filenames, uploaded_file_contents, n_clicks):
    """Save uploaded files and regenerate the file list."""
    global kmlname
    global name
    # Find which action has been changed/pressed.
    changed_id = [p['prop_id'] for p in callback_context.triggered][0]
    if "upload-data-RPL" in changed_id:
        if uploaded_filenames is not None and uploaded_file_contents is not None:
            name = ""
            for name, data in zip(uploaded_filenames, uploaded_file_contents):
                # Check file type is .xls and .xlsx
                if ".xls" in name or ".xlsx" in name:
                    save_file(name, data)
                else:
                    return {'display': 'none'}, "Incorrect file type. Only .xls and .xlsx is supported", ("", "")
            return {'display': 'none'}, name, ("", "")
        else:
            return {'display': 'none'}, "", ("", "")
    elif "RPL_to_KML_btn" in changed_id:
        rpl_list = get_RPL_file()
        try:
            # Load RPL into dataframe
            df = pd.read_excel(os.path.join(UPLOAD_DIRECTORY, rpl_list[0]))
        except NameError:
            return {'display': 'none'}, "RPL file already converted", ("", "")
        except IndexError:
            return {'display': 'none'}, "Please upload a RPL file", ("", "")
        # Save only longitude, latitude, cable type and slack columns
        df_cols = df.iloc[:, [2, 3, 5, 6, 23, 26]]
        # Drop first two rows containing header
        df_cols = df_cols.drop([0,1])
        # Drop NaN values by cells
        df2 = df_cols.apply(lambda x: pd.Series(x.dropna().values))

        # Remove file type extension
        if ".xlsx" in rpl_list[0]:
            RPL_name = rpl_list[0].replace(".xlsx", "")
        else:
            RPL_name = rpl_list[0].replace(".xls", "")

        # Initialise variables
        coords = []
        cable_type = []
        slack = []
        prev_cable_type = df2.iloc[:, 5].iloc[1]
        index = 0
        i = 0

        # Loop through each row of dataframe
        while index < df2.shape[0]:

            # Combine both longitude columns together.
            lat = (df2.iloc[:, 0].iloc[index]) + ((df2.iloc[:, 1].iloc[index]) / 60)
            # Combine both latitude columns together.
            lon = (df2.iloc[:, 2].iloc[index]) + ((df2.iloc[:, 3].iloc[index]) / 60)
            # Check if current cable type equals previous cable type
            # Different cable types indicate different sections
            if df2.iloc[:, 5].iloc[index] == prev_cable_type:
                try:
                    # Attempt to append coordinates into array using an index, indicating the same section.
                    slack[i].append(df2.iloc[:, 4].iloc[index])
                    cable_type[i].append(df2.iloc[:, 5].iloc[index])
                    coords[i].append([lon, lat])
                    prev_cable_type = df2.iloc[:, 5].iloc[index]

                except IndexError:
                    # If the section, index, has not been created yet, just add the coordinates, this creates the index, section.
                    slack.append([df2.iloc[:, 4].iloc[index]])
                    cable_type.append([df2.iloc[:, 5].iloc[index]])
                    coords.append([[lon, lat]])

            else:
                slack.append([df2.iloc[:, 4].iloc[index]])
                cable_type.append([df2.iloc[:, 5].iloc[index]])
                coords.append([[lon, lat]])
                prev_cable_type = df2.iloc[:, 5].iloc[index]
                i += 1
            index += 1

        # Create KML object
        kml = simplekml.Kml()

        n = 0
        while n <= len(coords):

            try:
                # For each set of coordinates in each section, different cable type, add a new linestring to the kml
                linestring = kml.newlinestring(name=RPL_name,coords=coords[n])
                # For each set of simpledata attributes in each section, add new simpledata field with corrospoding value
                linestring.extendeddata.schemadata.newsimpledata('cable_type', cable_type[n][0])
                linestring.extendeddata.schemadata.newsimpledata('slack_percent', slack[n][0])

            except IndexError:
                break
            # Increment i to cycle through coords and all the arrays
            n += 1

        # Remove chars from folder name to avoid error when creating a path and saving file
        RPL_name = RPL_name.replace('/', '')
        RPL_name = RPL_name.replace('\\', '')
        RPL_name = RPL_name.replace('&', '')
        RPL_name = RPL_name.replace('<', '')
        RPL_name = RPL_name.replace('>', '')
        RPL_name = RPL_name.replace('"', '')
        RPL_name = RPL_name.replace("'", '')

        # Set KML name and save directory
        kmlname = os.path.join(UPLOAD_DIRECTORY, (RPL_name + ".kml"))

        # Save KML
        kml.save(kmlname)

        # Remove RPL from UPLOAD_DIRECTORY
        for file in rpl_list:
            os.remove(os.path.join(UPLOAD_DIRECTORY, file))

        return {'display': 'inline-block'}, name, ("", "")

    else:
        return {'display': 'none'}, "", ("", "")
