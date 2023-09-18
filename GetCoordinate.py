import googlemaps
import pandas as pd

#Google Maps API key
api_key = 'API_KEY'

#Function to retrieve the latitude and longitude of an address.
def geocode_address(address, gmaps_client):
    try:
        geocode_result = gmaps_client.geocode(address)
        if geocode_result and len(geocode_result) > 0:
            location = geocode_result[0]['geometry']['location']
            return location['lat'], location['lng']
        else:
            return None, None
    except Exception as e:
        print(f"Error geocoding {address}: {str(e)}")
        return None, None

#Loading the spreadsheet
file_path = 'ExcelFile.xlsx'
data = pd.read_excel(file_path)

#Initializing the Google Maps client.
gmaps = googlemaps.Client(key=api_key)

#Lists to store latitudes and longitudes.
latitudes = []
longitudes = []

#Iterating over the addresses in the spreadsheet to obtain geographical coordinates
for index, row in data.iterrows():
    endereco = f"{row['Adress']}, {row['Neighborhood']}, {row['City']}"
    latitude, longitude = geocode_address(endereco, gmaps)
    latitudes.append(latitude)
    longitudes.append(longitude)

#Adding latitude and longitude columns to the DataFrame
data['Latitude'] = latitudes
data['Longitude'] = longitudes

#Saving the updated DataFrame to a new file.
output_file_path = 'address_spreadsheet_lat_long.xlsx'
data.to_excel(output_file_path, index=False)

print(f"Latitude and longitude have been added to the spreadsheet and saved in {output_file_path}.")
