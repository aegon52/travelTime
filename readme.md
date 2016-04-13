
# Distance Matrix api scraper documentation
Ben Norquist

## Dependencies
This script uses the following python libraries:
googlemaps, openpyxl,os,time,csv,datetime,math, and pprint

All are included in the Python Standard Library except for googlemaps and openpyxl. These libraries must be installed separately. The easiest way to do this is using pip. If you have pip preinstalled (standard for Python 2.7.9+) you can simply type ‘pip googlemaps’ to install the googlemaps library. 


## Inputs
*the script prompts the user for the following inputs*
The first prompt asks: **Origin -&gt; Destinations (one to many) (type "y") or Origins -&gt; Destination (many to one) (type "n")**
The first option will return travel times for auto and transit from one origin point to many destinations, the second option will return travel times from many origins to one destination. This is important for looking at peak hour travel patterns in areas with heavy commute flow in one direction. 

Origin or Destination point – Long/Lat or address (will be geocoded by google)
Departure time in the future – in ‘dd.mm.yyyy hh:mm:ss’ notation
Path to a .xlsx document containing the set of origins or destinations
To obtain input coordinates, first a bounding box was drawn in ArcGIS to the extent of the study area. Then, the "fishnet" tool was used to create an array of points at fixed interval (e.g. 1/2 mile) across the extent of the bounding box. Lat/Long values for the points created by the fishnet were used as inputs for the analysis. Simply save this data into a .xlsx file and note the file location. 
Header Row – this value accounts for any headers that you may have included in your input workbook. For example if you have one row of headers you would set this value = 1

Google API key – once you have activated the directions matrix api within your google developers console you will be given an API key (beginning with ‘aiza’) paste this value here

## Usage limits
The google directions matrix api allows 2,500 query elements/day for free, each OD pair uses 3 query elements as the script retrieves data for driving duration with the ‘best guess’ traffic model, driving duration with the pessimistic model and transit duration. So a set of 100 ODs would use 300 requests. 


## Output

The script writes the output data to a csv saved in the python working directory. The following data is recorded:

Driving distance
Driving duration (uncongested)
Driving duration in traffic ‘best guess’
Driving duration in traffic ‘pessimistic’
Transit distance
Transit duration
Difference (transit duration – driving duration (best guess))
Difference (transit duration – driving duration (pessimistic))

The output csv can be imported to ArcGIS using add data-&gt;xy data. Saving this as a shapefile and using symbology can produce heat maps of travel time as well as travel time ratio and absolute difference in travel time for each mode. 

