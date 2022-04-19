import pandas as pd
import yaml
from util import Sorting
import sys
import warnings

# if not sys.warnoptions:
warnings.filterwarnings('ignore')


if __name__ == "__main__":
    config_dict = {}
    with open('config.yaml') as f:
        data = yaml.load(f, Loader=yaml.FullLoader)
        config_dict = dict(data)
    sorting = Sorting(config_dict)
    sorting()
        