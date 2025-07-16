import pandas
import re
import os
import math
import pypdf
import datetime


def read_and_normalize(file_path) -> pandas.DataFrame:
    file = pandas.read_excel(file_path)
    headers_list = file.columns.tolist()
    found_headers_to_check = ['latitude', 'longitude', 'scid', 'node_type', 'ppl company_tag', 'pole_owner',
                              'verizon pennsylvania inc._tag', 'pole_tag', 'unknown_tag', 'make_ready_notes', 'address',
                              'commonwealth telephone co.  dba frontier comm._tag', 'county']
    found_headers = [header for header in found_headers_to_check if header in headers_list]
    file = file[found_headers]

    replacements = {
        'PPL Company': 'PPL',
        'Verizon Pennsylvania Inc.': 'Verizon',
        'Frontier Communications of PA. - New Holland': 'Frontier',
        'Frontier Communications of PA. - New Holland Telecom': 'Frontier',
        'Frontier Communications - Lakewood': 'Frontier',
        'Frontier Communications - Lakewood Telecom': 'Frontier',
        'Commonwealth Telephone Co.  dba Frontier Comm.': 'Frontier',
        'Commonwealth Telephone Co.  dba Frontier Comm. Telecom': 'Frontier',
        'Loop Telecom Pennsylvania LLC': 'Loop Internet',
        'UGI Utilities - Electric Division': 'UGI',
        'UGI Utilities - Gas': 'UGI',
        'UGI PENN NATURAL GAS, INC': 'UGI',
        'Service Electric Cablevision Inc - Mahanoy City': 'Service Electric',
        'Service Electric Cablevision': 'Service Electric',
        'Service Electric Cable TV Inc.': 'Service Electric',
        'Service Electric Company - Wilkes-Barre': 'Service Electric',
        'Upper Oxford Twp, Chester Co.': 'Xfinity',
        'City of Scranton - Wireless': 'City of Scranton',
        'City of Scranton': 'City of Scranton',
        'City of Scranton Office of Economic & Community Development': 'City of Scranton',
        'CTSI, LLC, dba Frontier Communications': 'CTSI'
    }

    # Modify owner column
    pole_owners = file['pole_owner'].tolist()
    for i, value in enumerate(pole_owners):
        if value in replacements:
            pole_owners[i] = replacements[value]
        else:
            pole_owners[i] = 'Surveyed / Unknown Owner'
    file['pole_owner'] = pole_owners

    file.rename(columns={
        'latitude': 'Latitude',
        'longitude': 'Longitude',
        'scid': 'SCID',
        'node_type': 'Pole Type',
        'ppl company_tag': 'Tag',
        'pole_owner': 'Owner',
        'make_ready_notes': 'Make Ready Notes'
    }, inplace=True)

    return file


def combine_tags(file) -> pandas.DataFrame:
    # Modify tag column
    tags_list = ['verizon pennsylvania inc._tag', 'pole_tag', 'unknown_tag']
    tag_variables = {}
    for tag in tags_list:
        if tag not in file.columns.tolist():
            tags_list.remove(tag)
        else:
            tag_variables[tag] = file[tag].tolist()
    tag_final = file['Tag'].tolist()
    for i, tag in enumerate(tag_final):
        if tag == '' or tag == 'NT':
            temp_comp = []
            for x in tag_variables:
                temp_comp.append(str(tag_variables[x][i]))
            tag_final[i] = max(temp_comp)
    file['Tag'] = tag_final
    return file


def vetro_export(file: pandas.DataFrame, path, name) -> bool:
    try:
        file = combine_tags(file)
        # Rename columns for Vetro import
        file[['Latitude', 'Longitude', 'SCID', 'Pole Type', 'Tag', 'Owner']].to_excel(f'{path}/{name}-Vetro-data.xlsx', index=False)
        return True
    except Exception as e:
        with open(f'{path}/error.log', 'a+') as f:
            f.write(str(e))
        return False


def generate_mrn(file: pandas.DataFrame, path, name) -> bool:
    try:
        file = combine_tags(file)
        # MRN Data Sheet
        file[['SCID', 'Tag', 'Latitude', 'Longitude', 'Make Ready Notes']].to_excel(f'{path}/{name}-MRN-data.xlsx', index=False)
        return True
    except Exception as e:
        with open(f'{path}/error.log', 'a+') as f:
            f.write(str(e))
        return False


def verizon_app(file: pandas.DataFrame, path, name) -> bool:

    def split_and_convert(target: str) -> int:
        if not target or not isinstance(target, str):
            raise ValueError("Input must be a non-empty string")
        try:
            parts = target.split('-')
            if len(parts) != 2:
                raise ValueError("Input must be in format 'feet-inches'")
            feet = int(parts[0])
            inches = int(parts[1])
            if inches < 0 or inches >= 12:
                raise ValueError("Inches must be between 0 and 11")
            if feet < 0:
                raise ValueError("Feet cannot be negative")
            return (feet * 12) + inches
        except ValueError:
            raise ValueError

    def get_street_name(address: str) -> str:
        return address.split(', ')[0].split(' ', maxsplit=1)[1]

    try:
        # Refactor file for only the relevant information
        file = file.loc[file['Owner'] == 'Verizon', ['Latitude', 'Longitude', 'SCID', 'Owner', 'Tag', 'verizon pennsylvania inc._tag',
                                                     'Make Ready Notes', 'address']]

        # Create a new dataframe
        columns_mrs = {
            'Pole Ref #': [],
            'Telco Pole #': [],
            'ELCO Pole #': [],
            'Route/Line (Verizon Use Only)': [],
            'Street Name': [],
            'Attacher Company': [],
            'Attachment Type': [],
            'Action': [],
            'Existing Height': [],
            'New Height': [],
            'Quantity': []
        }
        columns_info = {
            'Pole Ref #': [],
            'Telco Pole #': [],
            'ELCO Pole #': [],
            'Route/Line (Verizon Use Only)': [],
            'Street Name': [],
            'Attachment Description': [],
            'Number of Attachments': [],
            'Attachment Height': [],
            'Billing Description (Verizon Use Only)': [],
            'Fs/Rs OR Quad': [],
            'Comments': []
        }
        columns_details = {
            'Pole Ref #': [],
            'MR Req': [],
            'Telco Pole #': [],
            'ELCO Pole #': [],
            'Route/Line (Verizon Use Only)': [],
            'Street Name': [],
            'Cross Street Name': [],
            'Location Description': [],
            'Latitude': [],
            'Longitude': [],
            'Height': [],
            'Class': [],
            'Exclude from Application (Verizon Use Only)': [],
            'Not Owned or Controlled by VZ (Verizon Use Only)': [],
            'Customer Already Attached': [],
            'Pole OTMR Qualified Y/N (Verizon Use Only)': [],
            'If No, Reason Why (Verizon Use Only)': []
        }
        verizonmmrs = pandas.DataFrame(columns_mrs)
        verizoninfo = pandas.DataFrame(columns_info)
        verizondetails = pandas.DataFrame(columns_details)

        # Declare variables for information refactorization
        hardware = ['Guy', 'Com', 'Strand']
        action_check = ['Attach', 'Raise', 'Lower']
        companies = {
            'Verizon Pennsylvania Inc.': 'VERIZON WIRELESS(AERIAL)',
            'CTSI, LLC, Dba Frontier Communications': 'FRONTIER COMMUNICATIONS',
            'Loop Telecom Pennsylvania LLC': 'LOOP INTERNET HOLDCO LLC',
            'Comcast': 'COMCAST',
            'Service Electric Company': 'SERVICE ELECTRIC CABLE TV'
        }
        actions = {
            'Raise': 'Raise',
            'Lower': 'Lower',
            'Attach': 'No Make Ready'
        }
        attachments = {
            'Com': 'Cable/Strand',
            'Guy': 'Down Guy',
            'Strand': 'Cable/Strand',
        }
        power = ['Primary', 'Secondary', 'Drip Loop']
        file.reset_index(drop=True, inplace=True)
        mrn_iterable = file['Make Ready Notes'].tolist()

        # Declare variables for table input
        attacher_company = None
        attachment_type = None
        action = None
        existing_height = None
        new_height = None

        # Begin refactoring data
        for x, value in enumerate(mrn_iterable):
            if not file.loc[x, 'SCID'].isdigit():  # Skip pole if SCID contains a letter because it is a reference pole
                continue
            if pandas.isna(value) or not isinstance(value, str):  # Check for NaN/float values
                new_row_data = {
                    'Pole Ref #': file.loc[x, 'SCID'],
                    'Telco Pole #': file.loc[x, 'verizon pennsylvania inc._tag'],
                    'ELCO Pole #': file.loc[x, 'Tag'],
                    'Attacher Company': 'Not Surveyed',
                    'Attachment Type': 'n/a',
                    'Action': 'n/a',
                    'Existing Height': 'n/a',
                    'New Height': 'n/a',
                    'Quantity': 'n/a'
                }
                # Add the new row to the DataFrame using .loc
                verizonmmrs.loc[len(verizonmmrs)] = new_row_data
                continue  # Move to the next instruction
            lines = re.split('\n+', value)
            for line in lines:
                company_found = False
                for comp in companies:
                    comp_match_obj = re.match(comp, line, re.IGNORECASE)
                    if comp_match_obj:
                        attacher_company = companies[comp_match_obj.group()]
                        line = line[comp_match_obj.end():]
                        company_found = True
                        break  # Exit the loop once a company is found
                if not company_found:
                    continue  # Skip to the next line if no company was matched
                if not any(re.search(item, line, re.IGNORECASE) for item in hardware):
                    continue
                if not any(re.search(act, line, re.IGNORECASE) for act in action_check):
                    continue
                punctuation_marks = r'[:"\'.;]'
                line = re.sub(punctuation_marks, '', line)
                iterable_line = line.split(' ')
                for i, word in enumerate(iterable_line):
                    if word == 'at':
                        attachment_type = iterable_line[i - 1]
                        existing_height = iterable_line[i + 1]
                        action = iterable_line[i + 2]
                        # Perform calculations if needed for a new height
                        if action == 'Raise':
                            change = iterable_line[i + 3]
                            new_height = split_and_convert(existing_height) + int(change)
                            new_height = new_height / 12
                            fraction, whole = math.modf(new_height)
                            first, second = existing_height.split('-')
                            existing_height = f'''{first}'-{second}"'''
                            new_height = f'''{int(whole)}'-{int(fraction * 12)}"'''
                            break
                        elif action == 'Lower':
                            change = iterable_line[i + 3]
                            new_height = split_and_convert(existing_height) - int(change)
                            new_height = new_height / 12
                            fraction, whole = math.modf(new_height)
                            first, second = existing_height.split('-')
                            existing_height = f'''{first}'-{second}"'''
                            new_height = f'''{int(whole)}'-{int(fraction * 12)}"'''
                            break
                        else:
                            first, second = existing_height.split('-', maxsplit=1)
                            existing_height = f'''{first}'-{second}"'''
                            new_height = existing_height
                            break

                # Format attachment type
                attachment_type = attachments[attachment_type]
                # Format action
                action = actions[str(action)]

                new_row_data_mrs = {
                    'Pole Ref #': file.loc[x, 'SCID'],
                    'Telco Pole #': file.loc[x, 'verizon pennsylvania inc._tag'],
                    'ELCO Pole #': file.loc[x, 'Tag'],
                    'Attacher Company': attacher_company,
                    'Attachment Type': attachment_type,
                    'Action': action,
                    'Existing Height': existing_height,
                    'New Height': new_height,
                    'Quantity': '1'
                }
                verizonmmrs.loc[len(verizonmmrs)] = new_row_data_mrs

                new_row_data_info = {
                    'Pole Ref #': file.loc[x, 'SCID'],
                    'Telco Pole #': file.loc[x, 'verizon pennsylvania inc._tag'],
                    'ELCO Pole #': file.loc[x, 'Tag'],
                    'Attachment Description': attachment_type,
                    'Attachment Height': new_height
                }
                if attacher_company == 'LOOP INTERNET HOLDCO LLC':
                    verizoninfo.loc[len(verizoninfo)] = new_row_data_info

                new_row_data_details = {
                    'Pole Ref #': file.loc[x, 'SCID'],
                    'Telco Pole #': file.loc[x, 'verizon pennsylvania inc._tag'],
                    'ELCO Pole #': file.loc[x, 'Tag'],
                    'Street Name': get_street_name(file.loc[x, 'address']),
                    'Latitude': file.loc[x, 'Latitude'],
                    'Longitude': file.loc[x, 'Longitude']
                }
                if not verizondetails['Pole Ref #'].isin([file.loc[x, 'SCID']]).any() and attacher_company == 'LOOP INTERNET HOLDCO LLC':
                    verizondetails.loc[len(verizondetails)] = new_row_data_details

        with pandas.ExcelWriter(f'{path}/{name}-verizon-MRS.xlsx', engine='openpyxl') as writer:
            verizonmmrs.to_excel(writer, sheet_name='Make Ready', index=False)
            verizoninfo.to_excel(writer, sheet_name='Attachment Info', index=False)
            verizondetails.to_excel(writer, sheet_name='Pole Details', index=False)
        return True
    except Exception as e:
        with open(f'{path}/error.log', 'a+') as f:
            f.write(str(e))
        return False


def frontier_pdf(file: pandas.DataFrame, path, name) -> bool:

    def gen_table_data(parent_file: pandas.DataFrame, limit: int, start: int = 0) -> dict:
        all_pole_data = {}
        for i in range(limit):
            # Format address for route num and street name
            i += start
            address = parent_file.loc[i, 'address']
            address = address.split(', ')
            for part in address:
                if re.search(r'\d', part):
                    break
                else:
                    address.remove(part)
                    break
            municipality = address[1]
            address = address[0]
            address = address.split(' ', maxsplit=1)
            route, street = address[0], address[1]
            pole_data = {
                f'Telephone Co Pole Row{i + 1 - start}': parent_file.loc[i, 'commonwealth telephone co.  dba frontier comm._tag'],
                f'Power Pole Row{i + 1 - start}': 'dk',
                f'Street  LocationRow{i + 1 - start}': street,  # Two spaces here because Frontier is retarded
                f'Route NumberRow{i + 1 - start}': route,
                f'MunicipalityRow{i + 1 - start}': municipality,
                f'CountyRow{i + 1 - start}': parent_file.loc[i, 'county']
            }
            pole_identifier = {f'Pole ID {i}': pole_data}
            all_pole_data.update(pole_identifier)
        return all_pole_data

    file = file.loc[file['Owner'] == 'Frontier', ['Latitude', 'Longitude', 'SCID', 'Owner', 'Tag', 'Make Ready Notes', 'address', 'county', 'commonwealth telephone co.  dba frontier comm._tag']]
    file.reset_index(drop=True, inplace=True)
    date = datetime.datetime.now().strftime("%m/%d/%Y")
    poles_count = len(file['SCID'].tolist())

    # Create file structure and variable groups
    index, count = 0, 1
    groups = {}
    while poles_count > 0:
        poles_count -= 20
        groups[f'Group{count}'] = gen_table_data(file, poles_count + 20 if poles_count < 0 else 20, start=index)
        index += 20
        count += 1

    with pypdf.PdfReader('pdf/template.pdf') as reader:
        for group, poleID in groups.items():
            data_combined = {}
            for _, pole in poleID.items():
                data_combined.update(pole)
            for _, pole in poleID.items():
                reader_copy = reader  # Creates a copy of PDF to not overwrite the original PDF when new data is added
                output_path = f"{path}/{name}/{group}/{next(iter(pole.values()))}-frontier_form.pdf"
                writer = pypdf.PdfWriter()
                writer.clone_reader_document_root(reader_copy)
                writer.update_page_form_field_values(writer.pages[0], {'Date': date})
                os.makedirs(f"{path}/{name}/{group}", exist_ok=True)
                writer.update_page_form_field_values(writer.pages[0], data_combined)
                with open(output_path, "wb") as output_file:
                    writer.write(output_file)

    return True
