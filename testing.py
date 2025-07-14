import pypdf

# Make sure this points to the exact template you are using
pdf_template_path = 'pdf/template.pdf'

try:
    reader = pypdf.PdfReader(pdf_template_path)
    fields = reader.get_fields()

    if not fields:
        print("Error: No fillable form fields were found in this PDF.")
    else:
        print("--- Detailed PDF Field Inspection ---")
        for field_name, field_object in fields.items():
            print(f"\nPython Key: '{field_name}'")
            # .get_object() gives us the raw internal dictionary for the field
            print(f"  Internal PDF Object: {field_object.get_object()}")

except Exception as e:
    print(f"An error occurred: {e}")

# Note: @ should be replaced with the iterable number given by a for loop.
# Page 1 information

# Date
# Telephone Co Pole Row @
# Power Pole Row@
# Street LocationRow@
# Route NumberRow@  (This is just the house number)
# MunicipalityRow@
# CountyRow@





# Page 2 information (per pole)

# Company Type
# Power Pole Num
# Street Location
# CityBoro Township
# Field Personnel Name

# GPS Coordinates LatLong
# Telco Pole No.
# Name of Attacher
# Date

# Cable
# Cabinet
# Guy Pole
# Overlash

# Pole Size HtClass
# StreetlightYes
# StreetlightNo
# PowerConduitRiserYes
# PowerConduitRiserNo
# Trans_DeviceYes
# Trans_DeviceNo
# Street Light Bracket Height
# Top of Conduit Riser Height
# GuyingRequired
# GuyingNotRequired
# Attach Ht_@ -> max: 5
# Proposed
# MidSpan Ht_@ -> max: 10
# Span Length_@ -> max: 2
# Attacher Name_@ -> max: 4
# MRWYes
# MRWNo
# NESCYes
# NESCNo
# PoleHeightRequired
# PoleSide@ -> max: 5
