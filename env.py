import sys
from pathlib import Path

project_level_utils_path = Path(__file__).resolve().parent
global_level_utils_path = Path("/home/jovyan")

sys.path.insert(0, str(project_level_utils_path))
sys.path.insert(0, str(global_level_utils_path))

from utils.output_format import *

import pandas as pd
import numpy as np
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import OxmlElement
import re
import json
from docx.enum.table import WD_TABLE_ALIGNMENT
from scipy.special import erf
import scipy.stats as stats
from datetime import datetime
import yaml
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

config = "config.yaml"
with open(config, "r", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

raw_path = cfg["path"]["raw_path"]
pd_path = cfg["path"]["pd_path"]
code_path = cfg["path"]["code_path"]
remark_path = cfg["path"]["remark_path"]
timewin_path = cfg["path"]["timewin_path"]
output_path = cfg["path"]["output_path"]