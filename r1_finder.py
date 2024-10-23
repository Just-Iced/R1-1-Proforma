import openpyxl.workbook
import json, openpyxl, time
from bs4 import BeautifulSoup
from selenium import webdriver
from openpyxl import styles
from area import area
from turfpy.measurement import boolean_point_in_polygon
from geojson import Point, Polygon, Feature
from geopy.geocoders import Nominatim
# The class to find all heritage properties listed on popular realty websites
class R1_Finder:
    def __init__(self,
                    # The URL to retrieve the PDF from 
                    url: str = "https://guidelines.vancouver.ca/policy-vancouver-heritage-register.pdf",
                    # Where to store the data files
                    path: str = "data",
                    # Lines to ignore when reading the pdf
                    bad_lines: list[str] = ["8 HERITAGE BUILDINGS   ", "14 ARCHAEOLOGICAL SITES  ", ""]) -> None:
        
        self.url = url
        self.path = path
        self.bad_lines = bad_lines
        self.dictionary = {}
        self.parcels = {}
        try:
            self.dictionary = json.load(open(f"{path}/heritage_properties.json"))
        except FileNotFoundError:
            pass
        
        self.parcels = json.load(open("parcels.json"))
        
        
        firefox_profile = webdriver.FirefoxProfile()
        firefox_profile.set_preference("javascript.enabled", True)
        firefox_profile.set_preference("cookies.enabled", True)
        options = webdriver.FirefoxOptions()
        options.profile = firefox_profile
        self.driver = webdriver.Firefox(options=options)
        self.geolocator = Nominatim(user_agent="Your_Name")
        zoning = json.load(open("zoning.json"))
        self.zones = []
        for zone in zoning:
            polygon_list = [[]]
            for point in zone["coordinates"][0]:
                point.reverse()
                polygon_list[0].append(tuple(point))
            polygon = Polygon(polygon_list)
            self.zones.append(polygon)
        #self.vancouver_api_key = "f6aba1995ad5dcbd2f37c943d36001f1fcfa7441c6243868fa1a0792"
        #self.rentcast_key = "533d316d4fc54e3aac177815ae86d2f6"
            
    def get_sq_ft(self, address: str) -> float:
        address_data = {}
        try:
            address_data = self.parcels[address.strip()]
        except KeyError:
            return 0
        polygon = address_data["geom"]["geometry"]
        return round(area(polygon) * 10.764, 2)

    def is_r1_property(self, address: str) -> bool:
        coords = []
        try:
            raw_coords = self.parcels[address]["geo_point_2d"]
            print(raw_coords)
            coords = (float(raw_coords["lat"]), float(raw_coords["lon"]))
        except KeyError:
            location = self.geolocator.geocode(f"{address.strip()}, Vancouver", country_codes="CA", timeout=300, namedetails=True)
            if location == None:
                return False
            coords = (float(location.raw['lat']), float(location.raw['lon']))
        print(coords)
        point = Feature(geometry=Point(coords))
        
        
        for polygon in self.zones:
            if boolean_point_in_polygon(point, polygon):
                return True
        
        return False

    def format_realty_line(self, line: str, single_line: bool = False):
        line = line.title().strip()
        
        if "$" in line and not single_line:
            addresses = line.split("$")
            new_addresses = []
            for address in addresses[1:]:
                new_address = address.split("False")[1].split(", Vancouver")[0]
                new_line = self.format_realty_line(new_address, True)
                sqft = 0
                try:
                    sqft = self.get_sq_ft(new_line)
                except IndexError:
                    pass
                new_addresses.append((new_line, f"${address.split('False')[0].strip()}", "", sqft))
            return new_addresses
                
        else:
            line = line.removesuffix(", Vancouver, British Columbia")
            
            num_cnt = 0
            line_splt = line.split()
            first_word = True
            
            new_line = ""
            for word in line_splt:
                bad_word = False
                try:
                    int(word)
                    num_cnt += 1
                except ValueError:
                    if first_word:
                        bad_word = True
                first_word = False
                if not bad_word:
                    new_line += f"{word} "

            if num_cnt > 1:
                new_line = ""
                for word in line_splt[1:]:
                    new_line += f"{word} "
            if single_line:
                return f"{new_line}\n"
            else:
                with open(f"{self.path}/realty_listings_unformatted.txt", encoding="utf-8") as f:
                    cur_line = 0
                    lines = f.readlines()
                    for line in lines:
                        line = line.removesuffix("\n")
                        if new_line.lower().strip() in line.lower().strip():
                            break
                        cur_line += 1
                    
                    price = lines[cur_line - 2]
                    listing_time = lines[cur_line + 31]
                    sqft = self.get_sq_ft(new_line)
                    return f"{new_line}\n", price.removesuffix("\n").strip(), listing_time.removesuffix("\n").strip(), sqft
                    
    def format_realty_data(self) -> dict:
        realty_addresses_dict = {}
        with open(f"{self.path}/realty_listings_unformatted.txt", encoding="utf-8") as r:
            lines = r.readlines()
            for line in lines:
                if "vancouver" in line.lower():
                    formatted_line = self.format_realty_line(line)
                    if type(formatted_line) == list:
                        for address, price, list_time, sqft in formatted_line:
                            address = address.removesuffix("\n")
                            realty_addresses_dict[address] = {"Price": price, "Listing Time": list_time, "Sqft": float(sqft)}
                    else:
                        realty_addresses_dict[formatted_line[0]] = {"Price": formatted_line[1], "Listing Time": formatted_line[2], "Sqft": float(formatted_line[3])}
        return realty_addresses_dict
                    
    def update_realty_data(self) -> dict:
        def _goto_page(pg_num: int) -> None:
            url = f"https://www.realtor.ca/map#ZoomLevel=15&Center=49.2827%2C-123.1207&LatitudeMax=49.25428&LongitudeMax=-123.14113&LatitudeMin=49.19609&LongitudeMin=-123.22857&CurrentPage={pg_num}&Sort=6-D&PropertyTypeGroupID=1&TransactionTypeId=2&PropertySearchTypeId=0&Currency=CAD&HiddenListingIds=&IncludeHiddenListings=false"
            self.driver.get(url)
            self.driver.refresh()
        txt = ""

        _goto_page(1)
        time.sleep(20)
        soup = BeautifulSoup(self.driver.page_source, 'lxml')
        txt += f"{soup.get_text()}\n"
        time.sleep(3)
        for i in range(2, 51):
            print(f"Page {i}")
            _goto_page(i)
            time.sleep(3)
            soup = BeautifulSoup(self.driver.page_source, 'lxml')
            txt += f"{soup.get_text()}\n"
            time.sleep(3)
        
        with open(f"{self.path}/realty_listings_unformatted.txt", "w", encoding="utf-8") as f:
            f.write(txt)
        
        realty_addresses_dict = self.format_realty_data()
                    
        with open("data/realty_listings.txt", "w", encoding="utf-8") as f:
            f.writelines(realty_addresses_dict.keys())
            
        return realty_addresses_dict
    
    def get_r1_listings(self) -> None:
        realty_addresses_dict = self.format_realty_data()
        r1_addresses = []
        for address in realty_addresses_dict.keys():
            if self.is_r1_property(address.rstrip(" \n")):
                r1_addresses.append(realty_addresses_dict[address])
        self.dictionary = {}
        
        new_dict = {}
        for key, value in self.dictionary.items():
            if key not in new_dict.keys():
                new_dict[key.removesuffix(" \n").strip()] = value
        self.dictionary = new_dict
        with open(f"{self.path}/r1_1_properties.json", "w") as f:
            json.dump(self.dictionary, f, indent=4)
            
    def generate_spreadsheet(self) -> None:
        # Create the workbook
        wb = openpyxl.Workbook().file 

        # Get the active sheet and change its title
        sheet = wb.active
        sheet.title = "Heritage Property Profitability"

        sheet.cell(1, 1).value = "Address:"
        sheet.cell(1, 2).value = "Total Price:"
        sheet.cell(1, 3).value = "Size (ft²):"
        sheet.cell(1, 4).value = "Price per square foot:"
        sheet.column_dimensions['A'].width = 22
        sheet.column_dimensions['B'].width = 15
        sheet.column_dimensions['C'].width = 12
        sheet.column_dimensions['D'].width = 22
        

        list_num = 2
        for listing in self.dictionary:
            colour = None
            add_cell = sheet.cell(list_num, 1)
            add_cell.value = listing
            price_cell = sheet.cell(list_num, 2)
            sqft_cell = sheet.cell(list_num, 3)
            price_sqft_cell = sheet.cell(list_num, 4)
            sqft = self.dictionary[listing]["Sqft"]
            if sqft <= 0:
                price_sqft_cell.value = "Not Found"
                colour = styles.colors.Color(rgb='1F0000')
            else:
                price_sqft_cell.number_format = '[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00'
                price_cell.number_format = '[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00'
                price = float(self.dictionary[listing]['Price'].replace("$","").replace(",","").strip()) 
                sqft_cell.value = sqft
                price_cell.value = price
                price_sqft = round(price / sqft, 5)
                price_sqft_cell.value = price_sqft
                
                if self.is_profitable(price_sqft):
                    colour = styles.colors.Color(rgb='11FF00')
                else:
                    colour = styles.colors.Color(rgb='00FF0000')
                fill = styles.fills.PatternFill(patternType='solid', fgColor=colour)
                add_cell.fill = fill
                sqft_cell.fill = fill
                price_cell.fill = fill
                price_sqft_cell.fill = fill
                list_num += 1

        wb.save(f"{self.path}/Profitability.xlsx")
        
    def generate_proforma(self):
        from time import gmtime, strftime
        from openpyxl.worksheet.hyperlink import Hyperlink
        cur_day = strftime("%Y-%m-%d", gmtime())
        wb_title = f"Proforma {cur_day}.xlsx"
        wb = openpyxl.load_workbook("Proforma Template.xlsx")
        template = wb["Proforma Template"]
        master_list = wb["Master List"]
        from openpyxl.styles import Font
        format = Font(name="Arial", size=10)
        
        master_row = 5
        for key in self.dictionary:
            sheet = wb.copy_worksheet(template)
            data = self.dictionary[key]
            sheet.title = key
            sheet.cell(1,2).value = key
            sheet.cell(7,4).value = data["Sqft"]
            price = float(data['Price'].replace("$","").replace(",","").strip())
            sheet.cell(45,6).value = price
            
            
            address_cell = master_list.cell(master_row, 1)
            address_cell.value = key
            address_cell.hyperlink = Hyperlink("A1", f'{key}!A1', display=key)
            price_cell = master_list.cell(master_row, 2)
            price_cell.number_format = '[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00'
            price_cell.value = price
            
            
            master_row += 1
        wb.save(wb_title)
        wb.close()
        import xlwings 
        excel_app = xlwings.App(visible=False)
        excel_book = excel_app.books.open(wb_title)
        excel_book.save()
        excel_app.kill()
        
        wb = openpyxl.load_workbook(wb_title, data_only=True)
        for key in self.dictionary:
            sheet = wb[key]
            self.dictionary[key]["Profit"] = float(sheet.cell(94,7).internal_value)
        wb.close()
        wb = openpyxl.load_workbook(wb_title)
        master_row = 5
        master_list = wb["Master List"]
        for key in self.dictionary:
            profit_cell = master_list.cell(master_row, 3)
            profit_cell.number_format = "[$$-1009]#,##0.00;-[$$-1009]#,##0.00"
            profit = self.dictionary[key]["Profit"]
            profit_cell.value = profit
            if float(profit_cell.value) > 0:
                colour = styles.colors.Color(rgb='11FF00')
            else:
                colour = styles.colors.Color(rgb='00FF0000')
            fill = styles.fills.PatternFill(patternType='solid', fgColor=colour)
            profit_cell.fill = fill
            master_row += 1
        master_list.cell(4,2).font = format
        master_list.cell(4,3).font = format
        wb.remove(wb["Proforma Template"])
        wb.save(wb_title)

    #Convert the Vancouver parcel JSON into something more usable            
    def convert_parcel_json(self):
        try:
            with open("property-parcel-polygons.json") as f:
                new_dict = {}
                street_conversions = {
                    "St": "Street",
                    "Av": "Avenue",
                }
                
                data = json.load(f)
                for dictionary in data:
                    street_name = str(dictionary['streetname']).title()
                    splt = street_name.split()
                    if splt[-1] in street_conversions:
                        splt[-1] = street_conversions[splt[-1]]
                        street_name = ""
                        for word in splt:
                            street_name += f"{word} "
                    
                    new_dict[f"{dictionary['civic_number']} {street_name.strip()}"] = dictionary
                json.dump(new_dict, open("data/parcels.json", 'w'))
        except FileNotFoundError:
            return
        
    def convert_zoning_json(self):
        with open("zoning-districts-and-labels.json") as f:
            data = json.load(f)
            zoning_data = []
            for dictionary in data:
                if dictionary["zoning_district"] == "R1-1":
                    zoning_data.append(dictionary["geom"]["geometry"])
            json.dump(zoning_data, open("zoning.json", "w"))
        