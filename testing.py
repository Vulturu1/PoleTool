import sys
import os
import geopandas
import pyproj


def load_municipality_data(shp_file: str = None) -> dict:
    if not shp_file:
        try: source = sys._MEIPASS
        except: source =  os.path.abspath(".")
        source_path = os.path.join(source, 'vetro_source.zip')
    else:
        source_path = shp_file
    municipality_geo_data = {}
    datafromfile = geopandas.read_file(source_path)
    datafromfile = datafromfile[['MUNICIPA_1', 'geometry']]
    name = datafromfile['MUNICIPA_1'].tolist()
    geo = datafromfile['geometry'].tolist()
    in_proj = pyproj.Proj(init='epsg:3857')
    out_proj = pyproj.Proj(init='epsg:4326')

    for i in range(len(name)):
        if name[i] == 'None':
            continue
        municip = name[i]
        coord_temp = str(geo[i])
        coord_temp = coord_temp[coord_temp.rfind('(')+1:coord_temp.find(')', coord_temp.rfind('('))].split(', ')
        municipality_points = []
        print(municip)  # FIXME: REMOVE
        for c in coord_temp:
            print(f'working with {c}')
            lon, lat = c.split(' ')
            lon, lat = pyproj.transform(in_proj, out_proj, float(lon), float(lat))
            municipality_points.append((lon, lat))
        municipality_geo_data[municip] = municipality_points

    return municipality_geo_data
