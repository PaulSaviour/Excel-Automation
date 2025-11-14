fron werkzeug.utils import secure_filename

from flask import Flask, render template, request, redirect, url for, send file, session, flash

import os

import uvid

import json

import re

from pathlib import Path

from src.Helpers import main

import Api Call

import zipfile

import io

import threading

fron flask import jsonify

Import logging

import logging.config

import getpass

from datetime import datetime

import configparser

from dotenv import Load_dotenv

app Flask (_name__)

app.secret key = '46215442c98b1176996ee4ab24b6b5a1ecf8707cc37f118890b5a51d4e6a4d63'

cwd=os.getcwd()

project_root os.path.dirname(cwd)

input_folder os.path.join(project_root, 'input_folders')

input_folder2 os.path.join(project_root, 'validation_input')

#Save Invoice/PO contents here

#Save the single root-level .xlsx here

Path(input_folder).mkdir(parents=True, exist_ok=True)

Path(input_folder2).mkdir(parents=True, exist_ok=True)

input_folder3 os.path.join(project_root, 'src')

email_file_path = os.path.join(input_folder3, "email_id.txt")

report_dir_path = os.path.join(cwd, 'Output_File', 'Report_Files')

report_dir_path1 = os.path.join(cwd, 'Output_File', 'Data_Files')




Path(report_dir_path).mkdir(parents=True, exist_ok=True)

#pjvg

def init logging():

#Load.env so os.environ contains ENV

    load_dotenv()

    user name getpass.getuser()



    #BASE_LOG_DIR = fr"C:\Users\{user_name}\OneDrive WBA WBS\PSP\Capital_Project"

    one_drive_path = os.path.join(os.path.expanduser("-"), "OneDrive WBA", "WBS", "PSP", "Capital Project")

    BASE_LOG_DIR = one_drive_path
