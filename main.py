from r1_finder import R1_Finder

if __name__ == "__main__":
    # Create the property finder class
    propertyFinder = R1_Finder()
    # Make sure to be ready to do a captcha and click dismiss if you uncomment this line
    propertyFinder.get_r1_listings()
    propertyFinder.driver.close()
    propertyFinder.generate_proforma()
    