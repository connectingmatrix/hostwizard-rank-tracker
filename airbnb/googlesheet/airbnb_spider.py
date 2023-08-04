import re
import csv
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver


class AirbnbScraper:
    def __init__(self, location, checkin, checkout, guests, page_limit):
        self.location = location
        self.checkin = checkin
        self.checkout = checkout
        self.guests = guests
        self.driver = webdriver.Chrome()
        self.page_limit = page_limit
        self.url = ""

    def build_open_url(self):
        self.url = f"https://www.airbnb.com/s/{self.location}/homes?tab_id=home_tab&refinement_paths%5B%5D=%2Fhomes&flexible_trip_lengths%5B%5D=one_week&price_filter_input_type=0&price_filter_num_nights=1&date_picker_type=calendar&checkin={self.checkin}&checkout={self.checkout}&query={self.location}%2C%20FL&place_id=ChIJpzH_x2-p2YgRaLi3db561wc&source=structured_search_input_header&search_type=filter_change&adults={self.guests}"
        self.driver.get(self.url)

    def get_listings(self, search_page):
        self.driver.get(search_page)
        time.sleep(10)
        html = self.driver.page_source
        # Parse the HTML with BeautifulSoup
        soup = BeautifulSoup(html, 'html.parser')
        # Find all the listings on the page
        listings = soup.find_all('div', {'class': 'lwy0wad'})
        return listings

    def extract_element(self, listing_html, params):
        # 1. Find the right tag
        if 'class' in params:
            elements_found = listing_html.find_all(params['tag'], params['class'])
        else:
            elements_found = listing_html.find_all(params['tag'])

        # 2. Extract the right element
        tag_order = params.get('order', 0)
        element = elements_found[tag_order]

        # 3. Get text
        if 'get' in params:
            output = element.get(params['get'])
        else:
            output = element.get_text()

        return output

    def get_all_listings(self):
        all_listings = []
        get_url = self.url
        for i in range(self.page_limit):
            offset = 18 * i
            new_url = get_url + f'&items_offset={offset}&section_offset=3'
            new_listings = self.get_listings(new_url)
            all_listings.extend(new_listings)
            time.sleep(2)
        return all_listings

    def create_csv(self, scraped_data):
        df = pd.DataFrame(scraped_data)
        df['beds'] = df['rooms_and_beds'].str.extract(r'(\d+) beds')
        df['bedrooms'] = df['rooms_and_beds'].str.extract(r'(\d+) bedrooms')
        df['total_price'] = df['total_price'].str.replace('total', '')
        df['total_price'] = df['total_price'].str.extract(r'(\d+) $')
        df = df.drop('rooms_and_beds', axis=1)
        df = df.drop('price', axis=1)
        # Convert dataframe to CSV file
        df.to_csv('airbnb_data.csv', index=False)

    def scrape(self):
        RULES_SEARCH_PAGE = {
            'name': {'tag': 'div', 'class': 't1jojoys'},
            'description': {'tag': 'div', 'class': 'nquyp1l'},
            'rooms_and_beds': {'tag': 'div', 'class': 'f15liw5s', 'order': 1},
            'price': {'tag': 'div', 'class': '_1jo4hgw', },
            'total_price': {'tag': 'div', 'class': '_tt122m'},
        }

        all_listings = self.get_all_listings()
        scraped_data = []
        for i, listing_html in enumerate(all_listings):
            rank = i + 1
            name = self.extract_element(listing_html, RULES_SEARCH_PAGE['name'])
            description = self.extract_element(listing_html, RULES_SEARCH_PAGE['description'])
            rooms_and_beds = self.extract_element(listing_html, RULES_SEARCH_PAGE['rooms_and_beds'])
            price = self.extract_element(listing_html, RULES_SEARCH_PAGE['price'])
            total_price = self.extract_element(listing_html, RULES_SEARCH_PAGE['total_price'])
            data = {
                'rank': rank,
                'name': name,
                'description': description,
                'rooms_and_beds': rooms_and_beds,
                'price': price,
                'total_price': total_price,
            }
            scraped_data.append(data)
        self.create_csv(scraped_data)

    def main(self):
        self.build_open_url()
        self.scrape()


if __name__ == "__main__":
    location = "Hollywood,FL"
    checkin = "2023-03-02"
    checkout = "2023-03-03"
    guests = 12
    page_limit = 2
    spider = AirbnbScraper(location, checkin, checkout, guests, page_limit)
    spider.main()
