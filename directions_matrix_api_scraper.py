#dependencies
#must install googlemaps and openpyxl (available via pip), the remaining should be standar with your python installation
#Python version used for this script: 3.51

from pprint import pprint
import googlemaps, openpyxl,os,time,csv,datetime,math #progressbar (may add later)

###user input variables (can be hardcoded here by uncommenting to bypass user input step)
print('Please fill in the following required input variables: ')
#importPointsSheetLocation = 'path'
#o_d = True
o_dUserResponse = input('Origin -> Destinations (one to many) (type "y") or Origins -> Destination (many to one) (type "n")')
if 'y' in o_dUserResponse:
    o_d = True
elif 'n' in o_dUserResponse:
    o_d = False
else:
    print('invalid entry')
    quit()
#odPoint = 'some address/LongLat"
if o_d == True:
    odPoint = input('Origin Point? (address or long/lat)')
else:
    odPoint = input('Destination Point? (address or long/lat)')

importPointsSheetLocation = input("path to folder containing list of origins/destinations .xlsx file, be sure to include a / at the end  ")
#importPointsSheetLocation = '/file location path/'
#headerRow = 0
headerRow = input('header row? (input 0 if no headers  )')
headerRow = int(headerRow)
token = input('Google Maps API key  ')
# token = 'AIza...'
#departure_time = '10.03.2016 9:00:01'
oDsPerQuery = 50 ##origin destination points per query to gmaps api
departure_time = input('date and time in "dd.mm.yyyy hh:mm:ss" format  ')

print('input data: ','origin/destination response: ', o_dUserResponse, '\n o/d point: ',odPoint, '\n import excel sheet location: ',importPointsSheetLocation, '\n header row: ',headerRow, '\n departure time: ', departure_time)
###Functions

def open_file(path):
    #open and read ## add more here later to work with multiple input files
    files = os.listdir(path)
    paths = []
    for file in files:
        if '.xlsx' in file:
            paths.append(path + file)
    print('loading workbook...')
    wb = openpyxl.load_workbook(paths[0])
    print('wb loaded: ',paths[0])
    return wb

def impor_t(path,odPoint):
    wb = open_file(path)
    ws = wb.worksheets[0]
    longLatList = [odPoint]

    for row in ws.iter_rows(row_offset=headerRow):
        for cell in row:
            if cell.value is not None:
                longLatList.append(cell.value)
                print(cell.value)
    print(longLatList)
    return longLatList

def file_name(departure_time,o_d):
    name = departure_time[11:]
    if o_d ==True:
        name += '_O->Ds'
    elif o_d == False:
        name += '_Os->D'
    return name

def create_cutoffs(length, oDsPerQuery): #how many OD pairs are processes simultaneously
    chunks = length/oDsPerQuery
    chunk = math.ceil(chunks)
    cutOffs = [x*oDsPerQuery for x in range(1,chunk+1)]
    cutOffs[len(cutOffs)-1] = length
    print('cutoff points',cutOffs)
    return cutOffs

def web_request1(longLatList,token,depart_time): #for origins to destination
    ###loging to api and calculate epoch time
    gmaps = googlemaps.Client(key=token,timeout=10,queries_per_second=10)
    pattern = '%d.%m.%Y %H:%M:%S'
    epoch = int(time.mktime(time.strptime(depart_time, pattern)))
    print('epoch time converted:',epoch)
    resultDict = {'originLongLat':['origin','destination','driving_duration','duration_in_traffic','pessimistic_duration_in_traffic','distance ','transit_duration','transit_distance', 'diff','diff_pess']}
    resultCount = 0
    length = len(longLatList)-1
    cutOffs = create_cutoffs(length,oDsPerQuery)

    beginOfCutoff = 1
    print('cutoff val', longLatList[cutOffs[0]])
    for cutoff in cutOffs:
        uCutoff = cutoff #cutoff number for use in fetching

        #fetch batch of data
        carResult = gmaps.distance_matrix(longLatList[beginOfCutoff:uCutoff],longLatList[0],mode='driving',units='imperial',departure_time=epoch) ##change cutoffs[0] to cutoff to activate for loop
        pCarResult = gmaps.distance_matrix(longLatList[beginOfCutoff:uCutoff],longLatList[0],mode='driving',units='imperial',departure_time=epoch,traffic_model='pessimistic')
        transitResult = gmaps.distance_matrix(longLatList[beginOfCutoff:uCutoff],longLatList[0],mode='transit',units='imperial',departure_time=epoch)
        if beginOfCutoff == cutOffs[len(cutOffs)-1]:
            uCutoff = cutOffs[-1]

        pprint(carResult)
        for y in range(0,uCutoff-beginOfCutoff):
            if 'ZERO_RESULTS' in carResult.get('rows')[y].get('elements')[0].get('status'):
                print('not a routable address: ',longLatList[y+beginOfCutoff])
                break
            else:
                dest = carResult.get("destination_addresses")[0]
                orig = carResult.get("origin_addresses")[y]
                duration = carResult.get('rows')[y].get('elements')[0].get('duration').get('value')
                durationT = carResult.get('rows')[y].get('elements')[0].get('duration_in_traffic').get('value')
                dist = carResult.get('rows')[y].get('elements')[0].get('distance').get('value')
                durationPT = pCarResult.get('rows')[y].get('elements')[0].get('duration_in_traffic').get('value')
                #print('retrieving driving data ',longLatList[y])
            if 'ZERO_RESULTS' in transitResult.get('rows')[y].get('elements')[0].get('status'):
                tDuration = 'no transit route found'
                tDist = '-'
                #####pprint(transitResult) ## to print transit array
            else:
                tDuration = transitResult.get('rows')[y].get('elements')[0].get('duration').get('value')
                tDist = transitResult.get('rows')[y].get('elements')[0].get('distance').get('value')
            try:
                diff = float(tDuration)-float(durationT)
                diffPess = float(tDuration)-float(durationPT)
            except ValueError:
                diff = 'no transit route found'
                diffPess ='no transit route found'

            result = [longLatList[y+beginOfCutoff],orig,dest,duration,durationT,durationPT,dist,tDuration,tDist,diff,diffPess]
            #print('longLatList[y+beginOfCutoff]',longLatList[y+beginOfCutoff],duration,dist)
            resultCount += 1
            print("O/D's collected:",resultCount)

            ##fill dictionary of results
            key = result[0]
            lst = result[1:]
            resultDict[key] = lst
        beginOfCutoff = uCutoff

    ###create csv
    #pprint(resultDict)
    filename = 'd_matrix_output_' + file_name(depart_time) + '.csv'
    keys = sorted(resultDict.keys(),reverse=True)
    with open(filename,'w') as myfile:
        wr = csv.writer(myfile,delimiter = ',',quoting=csv.QUOTE_ALL)
        wr.writerow(keys)
        wr.writerows(zip(*[resultDict[key] for key in keys]))

    print("matched O/D's: ",len(resultDict),'resultCount: ', resultCount)

def web_request2(longLatList,token,depart_time): #for origin -> destinations
    ###loging to api and calculate epoch time
    gmaps = googlemaps.Client(key=token,timeout=10,queries_per_second=10)
    pattern = '%d.%m.%Y %H:%M:%S'
    epoch = int(time.mktime(time.strptime(depart_time, pattern)))
    print('epoch time converted:',epoch)
    resultDict = {'originLongLat':['origin','destination','driving_duration','duration_in_traffic','pessimistic_duration_in_traffic','distance ','transit_duration','transit_distance','diff','diff_pess']}
    resultCount = 0
    length = len(longLatList)-1
    cutOffs = create_cutoffs(length,oDsPerQuery)

    beginOfCutoff = 1
    print('cutoff val', longLatList[cutOffs[0]])
    for cutoff in cutOffs:
        uCutoff = cutoff #cutoff number for use in fetching

        #fetch batch of data
        carResult = gmaps.distance_matrix(longLatList[0],longLatList[beginOfCutoff:uCutoff],mode='driving',units='imperial',departure_time=epoch) ##change cutoffs[0] to cutoff to activate for loop
        pCarResult = gmaps.distance_matrix(longLatList[0],longLatList[beginOfCutoff:uCutoff],mode='driving',units='imperial',departure_time=epoch,traffic_model='pessimistic')
        transitResult = gmaps.distance_matrix(longLatList[0],longLatList[beginOfCutoff:uCutoff],mode='transit',units='imperial',departure_time=epoch)
        if beginOfCutoff == cutOffs[len(cutOffs)-1]:
            uCutoff = cutOffs[-1]

        pprint(carResult)
        for y in range(0,uCutoff-beginOfCutoff):
            if 'ZERO_RESULTS' in carResult.get('rows')[0].get('elements')[y].get('status'):
                print('not a routable address: ',longLatList[y+beginOfCutoff])
                break
            else:
                dest = carResult.get("destination_addresses")[y]
                orig = carResult.get("origin_addresses")[0]
                duration = carResult.get('rows')[0].get('elements')[y].get('duration').get('value')
                durationT = carResult.get('rows')[0].get('elements')[y].get('duration_in_traffic').get('value')
                dist = carResult.get('rows')[0].get('elements')[y].get('distance').get('value')
                durationPT = pCarResult.get('rows')[0].get('elements')[y].get('duration_in_traffic').get('value')
                #print('retrieving driving data ',longLatList[y])
            if 'ZERO_RESULTS' in transitResult.get('rows')[0].get('elements')[y].get('status'):
                tDuration = 'no transit route found'
                tDist = 'no transit route found'
                #####pprint(transitResult) ## to print transit array
            else:
                tDuration = transitResult.get('rows')[0].get('elements')[y].get('duration').get('value')
                tDist = transitResult.get('rows')[0].get('elements')[y].get('distance').get('value')

            try:
                diff = float(tDuration)-float(durationT)
                diffPess = float(tDuration)-float(durationPT)
            except ValueError:
                diff = 'no transit route found'
                diffPess = 'no transit route found'

            result = [longLatList[y+beginOfCutoff],orig,dest,duration,durationT,durationPT,dist,tDuration,tDist,diff,diffPess]
            #print('longLatList[y+beginOfCutoff]',longLatList[y+beginOfCutoff],duration,dist)
            resultCount += 1
            #print("O/D's collected:",resultCount)

            ##fill dictionary of results
            key = result[0]
            lst = result[1:]
            resultDict[key] = lst
        beginOfCutoff = uCutoff

    ###create csv
    #pprint(resultDict)
    filename = 'd_matrix_output_' + file_name(depart_time,o_d) + '.csv'
    keys = sorted(resultDict.keys(),reverse=True)
    with open(filename,'w') as myfile:
        wr = csv.writer(myfile,delimiter = ',',quoting=csv.QUOTE_ALL)
        wr.writerow(keys)
        wr.writerows(zip(*[resultDict[key] for key in keys]))

    print("matched O/D's: ",len(resultDict),'resultCount: ', resultCount)

############# execution

startTime = datetime.datetime.now()
print('start time: ',startTime)

list = impor_t(importPointsSheetLocation,odPoint)
print(list)
if o_d == True:
    web_request2(list,token,departure_time)
elif o_d == False:
    web_request1(list,token,departure_time)

endTime = datetime.datetime.now()
print('end time: ', endTime,'time elapsed: ',endTime-startTime)
