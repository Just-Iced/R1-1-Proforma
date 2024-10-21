from heritage_house_finder import HeritagePropertyFinder

if __name__ == "__main__":
    # Create the property finder class
    propertyFinder = HeritagePropertyFinder()
    propertyFinder.update_heritage_data()
    # Make sure to be ready to do a captcha and click dismiss if you uncomment this line
    propertyFinder.update_realty_data()
    propertyFinder.driver.close()
    propertyFinder.format_realty_data()
    propertyFinder.get_heritage_listings()
    propertyFinder.generate_proforma()
    
    