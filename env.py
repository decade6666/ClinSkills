import sys
from pathlib import Path

_env_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(_env_dir))

from utils.output_format import *  # noqa: F401,F403

import pandas as pd  # noqa: F401
import numpy as np  # noqa: F401
from docx import Document  # noqa: F401
from docx.shared import RGBColor  # noqa: F401
from docx.oxml.ns import nsdecls  # noqa: F401
from docx.oxml import OxmlElement  # noqa: F401
import re  # noqa: F401
import json  # noqa: F401
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: F401
from scipy.special import erf  # noqa: F401
import scipy.stats as stats  # noqa: F401
from datetime import datetime  # noqa: F401
import yaml  # noqa: F401
from openpyxl.utils.dataframe import dataframe_to_rows  # noqa: F401
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

config = _env_dir / "config.yaml"
with open(config, "r", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

raw_path = str(_env_dir / cfg["path"]["raw_path"])
pd_path = str(_env_dir / cfg["path"]["pd_path"])
code_path = str(_env_dir / cfg["path"]["code_path"])
remark_path = str(_env_dir / cfg["path"]["remark_path"])
timewin_path = str(_env_dir / cfg["path"]["timewin_path"])
output_path = str(_env_dir / cfg["path"]["output_path"])