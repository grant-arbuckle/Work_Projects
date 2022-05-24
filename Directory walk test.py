import pandas as pd
import os
from os.path import join, getsize

directory = '//in-fileserver/Business_Office/Month_End_Worksheets/POSTAGE/2022 Postage/'

for files in os.walk(directory):
  print(files)