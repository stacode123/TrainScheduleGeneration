#Import Libraries
import math
import warnings
from datetime import time
from os import makedirs as direct
import configparser

import re
import shutil

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Read configuration file
config = configparser.ConfigParser()
config.read('config.ini')
# Delete Existing Posters
shutil.rmtree(config['CONFIG']['OutputDirectory'], ignore_errors=True)


# Suppress all future warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
#Define Fonts
Nor= ImageFont.truetype(config['CONFIG']['FontPath'] , int(config['FONT_SIZES']['Normal']))
Norplus = ImageFont.truetype(config['CONFIG']['FontPath'] , int(config['FONT_SIZES']['NormalPlus']))
Bol= ImageFont.truetype(config['CONFIG']['BoldFontPath'] , int(config['FONT_SIZES']['Bold']))
Bols= ImageFont.truetype(config['CONFIG']['BoldFontPath'] , int(config['FONT_SIZES']['BoldSmall']))

def ReadExcel(sheet):
    df = pd.read_excel(config['CONFIG']['ExcelFilePath'], sheet_name=sheet)
    s=False
    index = df[(df.iloc[:, 0] == "Informacja o pociągu") | (df.iloc[:, 0] == "Train Info")].index[0]
    df = df.iloc[index:].reset_index(drop=True)
    s = False
    empty_row = pd.DataFrame([[None] * len(df.columns)], columns=df.columns)
    #df = pd.concat([empty_row, df], )
    index = df[(df.iloc[:, 0] == "Koniec") | (df.iloc[:, 0] == "End")].index[0]
    df = df.iloc[:index].reset_index(drop=True)
    df.fillna(method='ffill', inplace=True)
    return df







#Define Sort Keys
def sort_key(item):
    if item['departure_time'] == '?':
        return time.max  # Use the maximum time value for 'KOLIZJA'
    return item['departure_time']
def sort_key2(item):
    if isinstance(item, time):
        a = item.hour*60+item.minute
        return a
    return 0
def sort_key3(item):
    return item['departure_time']
def sort_key4(item):
    if isinstance(item, time):
        a = item.hour*60+item.minute
        if a <= 720:
            return a + 1440
        else:
            return a
    return 0




#Define Text Functions
def wrap_text(text, font, max_width):
    lines = []
    words = text.split()
    while words:
        line = ''
        while words and font.getlength(line + words[0]) <= max_width:
            line = line + (words.pop(0) + ' ')
        lines.append(line)
    return lines
def text_extract(list,station,current):
    textout = ""
    ind = list.index((station,current))
    for i in list[ind+1:]:
        textout = textout + i[0] + " " + str(i[1])[:5] + ", "
    return textout
#Define Text Cleaning Functions for Immersive Paintings
def remove_non_english_and_spaces(input_string):
    # Replace non-English letters and spaces with an empty string
    cleaned_string = re.sub(r'[^a-zA-Z]', '', input_string)
    return cleaned_string.lower()






#Open Excel File and set up variables
stations = []
sheet= pd.ExcelFile(config['CONFIG']['ExcelFilePath'])
sheets = sheet.sheet_names
per = len(sheets)
it = 0
print("Importing Data")
# Extract Station Names
for i in sheets:
    it+=1
    print(f"{(it / per) * 100:.2f}%")
    if "LK" in i:
        df = ReadExcel(i)
        df.fillna(method='ffill', inplace=True)
        column_a_data = df['Unnamed: 0'].tolist()
        stations = stations + column_a_data
# Remove station names that are not actual stations
stations = list(set(stations) - {" nan"})
stations = list(set(stations) - {"nan"})
stations = list(set(stations) - {"Informacja o pociągu"})
stations = list(set(stations) - {"Train Info"})
stations = list(set(stations) - {"Warszawa\xa0Zachodnia"})
# Populate Departure and Arrival Dictionaries
Departures =  {i: [] for i in stations}
Arrivals =  {i: [] for i in stations}

print("Parsing Data")
per = 2*len(sheets)
it = 0
# Extract Departure Times
for i in sheets:
    it+=1
    print(f"{(it / per) * 100:.2f}%")
    if "LK" in i:
        df = ReadExcel(i)
        df.fillna(method='ffill', inplace=True) # Drop rows where all elements are NaN

        train_details = df.iloc[:2].to_dict('records')
        for x in range(2,df.shape[1]):
            for index, row in df.iloc[2:].iterrows():
                station = row['Unnamed: 0']
                if station in Departures and row['Unnamed: 1'] == "odj.":
                    departure_time = row['Unnamed: {}'.format(x)]
                    if {'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]} not in Departures[station]:
                        if departure_time != '<' and departure_time != '|' and departure_time != '?' and departure_time == departure_time:
                           Departures[station].append({'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]})

# Extract Arrival Data
for i in sheets:
    it+=1
    print(f"{(it / per) * 100:.2f}%")
    if "LK" in i:
        df = ReadExcel(i)
        df.fillna(method='ffill', inplace=True) # Drop rows where all elements are NaN

        train_details = df.iloc[:2].to_dict('records')
        for x in range(2,df.shape[1]):
            for index, row in df.iloc[2:].iterrows():
                station = row['Unnamed: 0']
                if station in Arrivals and row['Unnamed: 1'] == "przyj." or row['Unnamed: 1'] == "przj." and station != "Warszawa\xa0Zachodnia":
                    departure_time = row['Unnamed: {}'.format(x)]
                    if {'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]} not in Arrivals[station]:
                        if departure_time != '<' and departure_time != '|' and departure_time == departure_time:
                            Arrivals[station].append({'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]})


trains = {}   #Dictionary of all stations and departure times for each train
trainsls = {} #Dictionary of all stations and arrival times for each train
trainslss = {} #Dictionary of the last station and arrival time for each train
#Generate train Dictionary
for i in Departures:
    for x in Departures[i]:
        trains[tuple(x['train_details'])] = []
        trainsls[tuple(x['train_details'])] = []
        trainslss[tuple(x['train_details'])] = []
#Sort the Dictionary
for key in trains:
    trains[key] = list(set(trains[key]))
for key in trainsls:
    trainsls[key] = list(set(trainsls[key]))
for key in trainslss:
    trainslss[key] = list(set(trainslss[key]))


#Itterate through all the stations and add the stations and departure times to the dictionary
for i in Departures:
    for x in Departures[i]:
        trains[tuple(x['train_details'])].append(i) # Append the station to the train
        trains[tuple(x['train_details'])].append(x['departure_time']) # Append the departure time to the train #
#Sort the dictionary by departure time
trainssort = {}
for key in trains:
    a = max([trains[key][i + 1] for i in range(0, len(trains[key]), 2)])
    b = min([trains[key][i + 1] for i in range(0, len(trains[key]), 2)])
    if sort_key2(a) - sort_key2(b) >= 720:
        trainssort[key] = sorted([(trains[key][i], trains[key][i + 1]) for i in range(0, len(trains[key]), 2)],key=lambda x: sort_key4(x[1]))
    else:
        trainssort[key] = sorted([(trains[key][i], trains[key][i + 1]) for i in range(0, len(trains[key]), 2)],key=lambda x: sort_key2(x[1]))

#Append Trains to trainsls
for i in Arrivals:
    for x in Arrivals[i]:
        if tuple(x['train_details']) not in trainsls:
            trainsls[tuple(x['train_details'])] = []
#Itterate through all the stations and add the stations and arrival times to the dictionary
for i in Arrivals:
    for x in Arrivals[i]:
        trainsls[tuple(x['train_details'])].append(i) # Append the station to the train
        trainsls[tuple(x['train_details'])].append(x['departure_time'])  # Append the departure time to the train

#Remove Duplicate Trains
for key in trainslss:
    trainslss[key] = list(set(trainslss[key]))

#Find Last Station for all trains and append to trainslss
for i in trainslss:
    a = max([trainsls[i][x + 1] for x in range(0, len(trainsls[i]), 2)])
    b = min([trainsls[i][x + 1] for x in range(0, len(trainsls[i]), 2)])
    if sort_key2(a) - sort_key2(b) >= 720:
        sort = (sorted(trainsls[i], key=sort_key4)[-1],trainsls[i][trainsls[i].index(sorted(trainsls[i], key=sort_key4)[-1])-1])
    else:
        sort = (sorted(trainsls[i], key=sort_key2)[-1],trainsls[i][trainsls[i].index(sorted(trainsls[i], key=sort_key2)[-1])-1])
    if trainslss[i] != sort:
        trainslss[i] = sort


print(trainslss)

#Make Sure The Directory Exists
direct(config['CONFIG']['OutputDirectory'], exist_ok=True)
#Create Posters
print("Creating Posters")
per = len(Departures)
it = 0
for i in Departures:
    it += 1
    print(f"{(it / per) * 100:.2f}%")
    iter=0
    page=0
    pos1 = (int(config['TEXT']['DepartureTimeX']), int(config['TEXT']['DepartureTimeY']))
    pos2 = (int(config['TEXT']['TrainNameX']), int(config['TEXT']['TrainNameY']))
    pos3 = (int(config['TEXT']['TrainNumberX']), int(config['TEXT']['TrainNumberY']))
    pos4 = (int(config['TEXT']['PassingStopsX']), int(config['TEXT']['PassingStopsY']))
    pos5 = (int(config['TEXT']['DestinationEndX']), int(config['TEXT']['DestinationEndY']))
    maxWidth = 400
    OdjSort =sorted(Departures[i], key=sort_key)
    image=Image.open(config['CONFIG']['BaseImage'])
    draw = ImageDraw.Draw(image)
    draw.text((int(config['TEXT']['StationNameX']), int(config['TEXT']['StationNameY'])), str(i), font=Bol, fill="black")
    for x in range(10*math.floor(len(Departures[i])/10)):
        iter+=1
        # Print Departure Time and Train Details
        draw.text(pos1, str(OdjSort[iter - 1]['departure_time'])[:5], font=Bols,fill="black")
        draw.text(pos2,str(list(OdjSort[iter-1]['train_details'])[0]), font=Nor, fill="black")
        draw.text(pos3, str(list(OdjSort[iter-1]['train_details'])[1]), font=Nor, fill="black")

        #Print Route on Page (Wrapped)
        wrapped_text = wrap_text(text_extract(trainssort[tuple(OdjSort[iter-1]['train_details'])],i,OdjSort[iter-1]['departure_time']), Norplus, maxWidth)
        y_text = pos4[1]
        for line in wrapped_text:
            draw.text((pos4[0],y_text),line, font=Norplus, fill="black")
            y_text += 13
        #Print Destination
        dest = trainslss[tuple(OdjSort[iter - 1]['train_details'])]
        text = str(dest[0])[:5] + " " + str(dest[1])
        draw.text((pos5[0]-Bol.getlength(text),pos5[1]),text, font=Bol, fill="black")


        #iter positions
        pos1 = (int(config['TEXT']['DepartureTimeX']),pos1[1]+int(config['TEXT']['IncrementY']))
        pos2 = (int(config['TEXT']['TrainNameX']),pos2[1]+int(config['TEXT']['IncrementY']))
        pos3 = (int(config['TEXT']['TrainNumberX']),pos3[1]+int(config['TEXT']['IncrementY']))
        pos4 = (int(config['TEXT']['PassingStopsX']),pos4[1]+int(config['TEXT']['IncrementY']))
        pos5 = (int(config['TEXT']['DestinationEndX']),pos5[1]+int(config['TEXT']['IncrementY']))
        #Start new page every 10 entries
        if iter%10==0:
            page+=1
            original_width, original_height = image.size
            res = image.resize((359, 507))
            new_image = Image.new("RGBA", (original_width, original_height), (255, 255, 255, 0))
            paste_x = 0
            paste_y = (original_height - 507) // 2
            new_image.paste(res, (paste_x, paste_y))

            new_image.save(f"{config['CONFIG']['OutputDirectory']}/{remove_non_english_and_spaces(i)}{page}.png")
            image=Image.open(config['CONFIG']['BaseImage'])
            draw = ImageDraw.Draw(image)
            draw.text((55, 37), str(i), font=Bol, fill="black")
            pos1 = (int(config['TEXT']['DepartureTimeX']), int(config['TEXT']['DepartureTimeY']))
            pos2 = (int(config['TEXT']['TrainNameX']), int(config['TEXT']['TrainNameY']))
            pos3 = (int(config['TEXT']['TrainNumberX']), int(config['TEXT']['TrainNumberY']))
            pos4 = (int(config['TEXT']['PassingStopsX']), int(config['TEXT']['PassingStopsY']))
            pos5 = (int(config['TEXT']['DestinationEndX']), int(config['TEXT']['DestinationEndY']))
    page+=1
    #Start last Page
    for x in range(len(Departures[i])%10):
        iter += 1
        # Print Departure Time and Train Details
        draw.text(pos1, str(OdjSort[iter - 1]['departure_time'])[:5], font=Bols, fill="black")
        draw.text(pos2, str(list(OdjSort[iter - 1]['train_details'])[0]), font=Nor, fill="black")
        draw.text(pos3, str(list(OdjSort[iter - 1]['train_details'])[1]), font=Nor, fill="black")

        # Print Route on Page (Wrapped)
        wrapped_text = wrap_text(text_extract(trainssort[tuple(OdjSort[iter-1]['train_details'])],i,OdjSort[iter-1]['departure_time']), Norplus, maxWidth)
        y_text = pos4[1]
        for line in wrapped_text:
            draw.text((pos4[0], y_text), line, font=Norplus, fill="black")
            y_text += 13
        #Print Destination
        dest = trainslss[tuple(OdjSort[iter - 1]['train_details'])]
        text = str(dest[0])[:5] + " " + str(dest[1])
        draw.text((pos5[0]-Bol.getlength(text),pos5[1]),text, font=Bol, fill="black")


        # iter positions
        pos1 = (int(config['TEXT']['DepartureTimeX']),pos1[1]+int(config['TEXT']['IncrementY']))
        pos2 = (int(config['TEXT']['TrainNameX']),pos2[1]+int(config['TEXT']['IncrementY']))
        pos3 = (int(config['TEXT']['TrainNumberX']),pos3[1]+int(config['TEXT']['IncrementY']))
        pos4 = (int(config['TEXT']['PassingStopsX']),pos4[1]+int(config['TEXT']['IncrementY']))
        pos5 = (int(config['TEXT']['DestinationEndX']),pos5[1]+int(config['TEXT']['IncrementY']))
    #Save Final Image For Immersive Painting Importing
    original_width, original_height = image.size
    res = image.resize((359, 507))
    new_image = Image.new("RGBA", (original_width, original_height), (255, 255, 255, 0))
    paste_x = 0
    paste_y = (original_height - 507) // 2
    new_image.paste(res, (paste_x, paste_y))

    new_image.save(f"{config['CONFIG']['OutputDirectory']}/{remove_non_english_and_spaces(i)}{page}.png")



