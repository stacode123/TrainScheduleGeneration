import math
import warnings
from datetime import time

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Suppress all future warnings
warnings.simplefilter(action='ignore', category=FutureWarning)



Nor= ImageFont.truetype("arial.ttf", 10)
Norplus = ImageFont.truetype("arial.ttf", 11)
Bol= ImageFont.truetype("arialbd.ttf", 18)
Bols= ImageFont.truetype("arialbd.ttf", 12)
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




stacje = []
sheet= pd.ExcelFile('rj.xlsx')
sheets = sheet.sheet_names
per = len(sheets)
it = 0
print("Importing Data")
for i in sheets:
    it+=1
    print(f"{(it / per) * 100:.2f}%")
    if "LK" in i:
  #      print("huj")
        df = pd.read_excel('rj.xlsx', sheet_name=i,)
        df.fillna(method='ffill', inplace=True)
   #     print(df)
        column_a_data = df['Unnamed: 0'].tolist()
    #    print(column_a_data)
        stacje = stacje + column_a_data

stacje = list(set(stacje) - {" nan"})
stacje = list(set(stacje) - {"nan"})
stacje = list(set(stacje) - {"Informacja o pociągu"})
stacje = list(set(stacje) - {"Warszawa\xa0Zachodnia"})
odjazdy =  {i: [] for i in stacje}

print("Parsing Data")
per = 2*len(sheets)
it = 0
# Iterate over each sheet again to add departure time and train details
for i in sheets:
    it+=1
    print(f"{(it / per) * 100:.2f}%")
    if "LK" in i:
        df = pd.read_excel('rj.xlsx', sheet_name=i)
        df.fillna(method='ffill', inplace=True) # Drop rows where all elements are NaN

        train_details = df.iloc[:2].to_dict('records')
        for x in range(2,df.shape[1]):
            for index, row in df.iloc[2:].iterrows():
                station = row['Unnamed: 0']
                if station in odjazdy and row['Unnamed: 1'] == "odj.":
                    departure_time = row['Unnamed: {}'.format(x)]
                    if {'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]} not in odjazdy[station]:
                        if departure_time != '<' and departure_time != '|' and departure_time != '?':
                           odjazdy[station].append({'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]})
przyjazdy =  {i: [] for i in stacje}
for i in sheets:
    it+=1
    print(f"{(it / per) * 100:.2f}%")
    if "LK" in i:
        df = pd.read_excel('rj.xlsx', sheet_name=i)
        df.fillna(method='ffill', inplace=True) # Drop rows where all elements are NaN

        train_details = df.iloc[:2].to_dict('records')
        for x in range(2,df.shape[1]):
            for index, row in df.iloc[2:].iterrows():
                station = row['Unnamed: 0']
                if station in przyjazdy and row['Unnamed: 1'] == "przyj." or row['Unnamed: 1'] == "przj." and station != "Warszawa\xa0Zachodnia":
                    departure_time = row['Unnamed: {}'.format(x)]
                    if {'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]} not in przyjazdy[station]:
                        if departure_time != '<' and departure_time != '|':
                            przyjazdy[station].append({'departure_time': departure_time,'train_details': [train_details[0]["Unnamed: {}".format(x)],train_details[1]["Unnamed: {}".format(x)]]})


trains = {}   #Dictionary of all stations and departure times for each train
trainsls = {} #Dictionary of all stations and arrival times for each train
trainslss = {} #Dictionary of the last station and arrival time for each train
#Generate train Dictionary
for i in odjazdy:
    for x in odjazdy[i]:
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
for i in odjazdy:
    for x in odjazdy[i]:
        trains[tuple(x['train_details'])].append(i) # Append the station to the train
        trains[tuple(x['train_details'])].append(x['departure_time']) # Append the departure time to the train #
#Sort the dictionary by departure time
trainssort = {}
for key in trains:
    trainssort[key] = sorted([(trains[key][i], trains[key][i + 1]) for i in range(0, len(trains[key]), 2)],key=lambda x: sort_key2(x[1]))

#Append Trains to trainsls
for i in przyjazdy:
    for x in przyjazdy[i]:
        if tuple(x['train_details']) not in trainsls:
            trainsls[tuple(x['train_details'])] = []
#Itterate through all the stations and add the stations and arrival times to the dictionary
for i in przyjazdy:
    for x in przyjazdy[i]:
        trainsls[tuple(x['train_details'])].append(i) # Append the station to the train
        trainsls[tuple(x['train_details'])].append(x['departure_time'])  # Append the departure time to the train

#Remove Duplicate Trains
for key in trainslss:
    trainslss[key] = list(set(trainslss[key]))

#Find Last Station for all trains and append to trainslss
for i in trainslss:
    sort = (sorted(trainsls[i], key=sort_key2)[-1],trainsls[i][trainsls[i].index(sorted(trainsls[i], key=sort_key2)[-1])-1])
    if trainslss[i] != sort:
        trainslss[i] = sort


print(trainslss)



#Create Posters
print("Creating Posters")
per = len(odjazdy)
it = 0
for i in odjazdy:
    it += 1
    print(f"{(it / per) * 100:.2f}%")
    iter=0
    page=0
    pos1 = (54,104)
    pos2 = (92,100)
    pos3 = (94,110)
    pos4 = (160,102)
    pos5 = (560,135)
    maxWidth = 400
    OdjSort =sorted(odjazdy[i], key=sort_key)
    image=Image.open("Podstawa.png")
    draw = ImageDraw.Draw(image)
    draw.text((55, 37), str(i), font=Bol, fill="black")
    for x in range(10*math.floor(len(odjazdy[i])/10)):
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
        pos1 = (54,pos1[1]+69)
        pos2 = (92,pos2[1]+69)
        pos3 = (94,pos3[1]+69)
        pos4 = (160,pos4[1]+69)
        pos5 = (560,pos5[1]+69)
        #Start new page every 10 entries
        if iter%10==0:
            page+=1
            image.save(f"Plakaty/{i}{page}.png")
            image=Image.open("Podstawa.png")
            draw = ImageDraw.Draw(image)
            draw.text((55, 37), str(i), font=Bol, fill="black")
            pos1 = (54,104)
            pos2 = (92,100)
            pos3 = (94,110)
            pos4 = (160,102)
            pos5 = (560,135)
    page+=1
    #Start last Page
    for x in range(len(odjazdy[i])%10):
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
        pos1 = (54, pos1[1] + 69)
        pos2 = (92, pos2[1] + 69)
        pos3 = (94, pos3[1] + 69)
        pos4 = (160,pos4[1] + 69)
        pos5 = (560,pos5[1] + 69)
    #Save Final Image
    image.save(f"Plakaty/{i}{page}.png")



