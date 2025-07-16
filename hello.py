
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut

locator = Nominatim(user_agent="my_app_name")
addresses = ["Scranton, PA", "Wilkes-Barre, PA", "Philadelphia, PA"]

for address in addresses:
    try:
        location = locator.geocode(address)
        if location:
            print(f"{address}: {location.latitude}, {location.longitude}")
        else:
            print(f"{address}: Not found")

    except GeocoderTimedOut:
        print(f"Error: Geocoder timed out for {address}")
