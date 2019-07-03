#!/usr/bin/env python3

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib
import pandas as pd


# =============================================================================
#               Define filepaths and read data from file
# =============================================================================
# Define filepath for import and export and filename for import file
path = r'\\filsrv01\BKI\11. Økonomi\04 - Controlling\NMO\1. Produktion\Procestegninger, diagrammer o.l\CoffeeWheel'
dataFile = r'\CoffeWheelData.xlsx'
destPath = r'\\filsrv01\BKI\11. Økonomi\04 - Controlling\NMO\1. Produktion\Procestegninger, diagrammer o.l\CoffeeWheel\Charts'
# Read data from Excel file
df = pd.read_excel(path + dataFile, sheet_name='Data')
# =============================================================================
#                           Define functions
# =============================================================================
# Function to translate value (0//1) from file to color in chart
def plot_color(dictionary, values):
    return [dictionary.get(item, item) for item in values]


# Function to turn dataframe into nested dictionary
def dataframe_to_dictionary(frame):
    dic = {}
    for row in frame.values:
        here = dic
        for elem in row[:-2]:
            if elem not in here:
                here[elem] = {}
            here = here[elem]
        here[row[-2]] = row[-1]
    return dic
# =============================================================================
#                       Define dictionaries and lists
# =============================================================================
# Define dictionaries with color per value
harvest = {0: plt.cm.Greys(0.0), 1: plt.cm.Greens(0.7)}
shipping = {0: plt.cm.Greys(0.0), 1: plt.cm.Blues(0.7)}
order = {0: plt.cm.Greys(0.0), 1: plt.cm.Reds(0.7)}
# Define layout for chart title
fontTitle = {'family': 'serif', 'color': 'black', 'weight': 'bold', 'size': 40
             , 'verticalalignment': 'top', 'horizontalalignment': 'center'}
# Define months and chart block-sizes
months = ('Jan', 'Feb', 'Mar', 'Apr', 'Maj', 'Jun', 'Jul', 'Aug'
          , 'Sep', 'Okt', 'Nov', 'Dec')
groupSize = (2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
# Convert dataframe to dictionary
wheelDic = dataframe_to_dictionary(df)
# Get unique country values from dataframe
countries = df.Country.unique()
# =============================================================================
#                            Define chart settings
# =============================================================================
# Font size for text not in title
matplotlib.rcParams['font.size'] = 13
# Define legends for charts
labelHarvest = mpatches.Patch(color=plt.cm.Greens(0.7), label='Harvest')
labelShipping = mpatches.Patch(color=plt.cm.Blues(0.7), label='Shipping')
labelOrder = mpatches.Patch(color=plt.cm.Reds(0.7), label='Order')
labels = [labelHarvest, labelShipping, labelOrder]
# =============================================================================
#                           Create chart per country
# =============================================================================
for country in countries:
# Outer ring (Harvest)
    fig, ax = plt.subplots()
    ax.axis('equal')
    harvestPie, _ = ax.pie(groupSize, counterclock=False, radius=3.0
                           , labels=months, startangle=-270
                           , colors=plot_color(harvest, [wheelDic[country]['Harvest'].get(val) for val in months]))
    plt.setp(harvestPie, width=1.0, edgecolor='black')
# Second ring (Shipping)
    shippingPie, _ = ax.pie(groupSize, counterclock=False, radius=3.0-1.0
                            , labels=None, startangle=-270
                            , colors=plot_color(shipping, [wheelDic[country]['Shipping'].get(val) for val in months]))
    plt.setp(shippingPie, width=1.0, edgecolor='black')
# Third ring (Order)
    orderPie, _ = ax.pie(groupSize, counterclock=False, radius=3.0-1.0-1.0
                         , labels=None, startangle=-270
                         , colors=plot_color(order, [wheelDic[country]['Order'].get(val) for val in months]))
    plt.setp(orderPie, width=1.0, edgecolor='black')
# Insert chart title and save charts to location
    plt.text(0, 4.0, country, fontdict=fontTitle)
    plt.legend(bbox_to_anchor=(1.2, 1.5), loc=3, handles=labels, prop={'size': 12})
    plt.savefig(destPath + '\\' + country + '.png', bbox_inches='tight')
