from dash import html
import base64
import os
from os.path import basename
from urllib.parse import quote as urlquote
from bs4 import BeautifulSoup
from flask import Flask, send_from_directory
from dash import dcc, html, Input, Output
import simplekml
from concurrent.futures import ThreadPoolExecutor
from zipfile import ZipFile
from RPLTool import app, server, quit1
import dash_bootstrap_components as dbc
# Run this app with `python app.py` and
# visit http://127.0.0.1:8050/ in your web browser.
from pathlib import Path
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
import xlwings as xw
import simplekml
from concurrent.futures import ThreadPoolExecutor
from zipfile import ZipFile
import time

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

        # Save only longitude and latitude columns
        df_cols = df.iloc[:, [2, 3, 5, 6]]
        # Drop and delete NaN rows
        df2 = df_cols.dropna()

        # Remove file type extension
        if ".xlsx" in rpl_list[0]:
            RPL_name = rpl_list[0].replace(".xlsx", "")
        else:
            RPL_name = rpl_list[0].replace(".xls", "")

        coords = []
        index = 0
        # Loop through each row of dataframe
        while index <= len(df.index):
            try:
                # Combine both longitude columns together.
                lat = (df2.iloc[:, 0].iloc[index]) + ((df2.iloc[:, 1].iloc[index]) / 60)
                # Combine both latitude columns together.
                lon = (df2.iloc[:, 2].iloc[index]) + ((df2.iloc[:, 3].iloc[index]) / 60)
                # Append coordinate to coords array
                coords.append([lon, lat])
            except IndexError:
                pass
            index += 1

        # Create KML object
        kml = simplekml.Kml()
        # Add one linestring as no section data in RPL
        kml.newlinestring(name=RPL_name, coords=coords)
        # Remove slashes from folder name to avoid error when creating a path and saving file
        RPL_name = RPL_name.replace('/', '')
        RPL_name = RPL_name.replace('\\', '')

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
