import requests
from bs4 import BeautifulSoup

import os
import re

# Group1: 107, 106
# Group2: 

url_prefix = "https://dep.mohw.gov.tw/DOS/"

def downloadAllXls(data, year):
	print(year)
	year_dir = os.getcwd() + "/data/" + str(year) + "/"
	for d in data:
		sheet_html = requests.get(url_prefix + d.find('a')['href'])
		soup = BeautifulSoup(sheet_html.text, 'lxml')
		sheets = soup.find("section", class_="listTb").findAll("a")
		for sheet in sheets:
			if "xls" in sheet["title"].lower():
				response = requests.get(sheet['href'])
				with open(year_dir + sheet['title'].split("(.")[0]+".xls", "wb") as f:
					f.write(response.content)


def parse(year_url, year):
	year_html = requests.get(url_prefix + year_url)
	soup = BeautifulSoup(year_html.text, 'lxml')
	data = soup.find("section", class_="nplist").findAll("li")
	if year <= 105 and year >= 103:
		for d in data:
			if "統計表" in d.text:
				data_url = url_prefix + d.find('a')['href']
		table_html = requests.get(data_url)
		table_soup = BeautifulSoup(table_html.text, 'lxml')
		table = table_soup.find("section", class_="nplist").findAll("li")
		data = table
	
	downloadAllXls(data, year)
		
def main():
	curr_dir = os.getcwd() + "/data/"
	if not os.path.isdir(curr_dir):
		os.mkdir(curr_dir)
	html = requests.get(url_prefix+"np-1918-113.html")
	soup = BeautifulSoup(html.text, 'lxml')
	years_report = soup.find("section", class_="nplist").findAll("li")
	
	for year_report in years_report:
		# print(year.find('a')['href'], year.text)
		year = re.search(r'\d*', year_report.text).group()
		if int(year) >= 106:
			continue
		year_url = year_report.find('a')['href']
		if not os.path.isdir(curr_dir+year):
			os.mkdir(curr_dir+year)
		else:
			for f in os.listdir(curr_dir+year):
				os.remove(os.path.join(curr_dir,year,f)) 
		parse(year_url, int(year))

if __name__ == "__main__":
	main()