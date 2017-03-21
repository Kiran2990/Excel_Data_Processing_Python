# Install required files before running code 
# pip install xlrd
# pip install xlwt

# Python version 2.7

# This code takes the name of the city and displays the closest city reachable from that city 
# Before running the code please make sure the Excel file and the program file are in the same folder
# running the code : python distance-finder.py

import xlrd

try:
	# This handles the opening of the workbook
	workbook = xlrd.open_workbook('cityData.xlsx')
	worksheet = workbook.sheet_by_index(0)
	num_of_rows = worksheet.nrows	
	num_of_columns = worksheet.ncols

	city_name = {}

	# This handles the building of the dictonary
	# We are building a dictonary of dictonaries
	# Each City is the key of the external dictonary and the value for these keys 
	# are the distance from the current city to the given cities
	# Example : {'Mumbai' : {'Thane': 22.7, 'Hyderabad': 738.0, 'Pune': 148.0 ...}}
	for row in range(1,num_of_rows):
		dist_from_city = {}
		for col in range(1,num_of_columns):
			if not row==col:
				dist_from_city.update({worksheet.cell(0, col).value : worksheet.cell(row, col).value})
		city_name.update({worksheet.cell(row, 0).value : dist_from_city})

	entered_city_name = raw_input('Enter city name: ')

	try:
		# This handles the sorting of elements in the dictonary
		#dist_from_city = city_name[entered_city_name]
		#sorted_dict = sorted(dist_from_city.items(),key=lambda x: x[1])
		#for i in range(len(sorted_dict)):
		#	print i+1,'.',sorted_dict[i][0] ,':', sorted_dict[i][1] , 'miles'

		travel_path = entered_city_name
		closest_city= entered_city_name
		visited_cities = [entered_city_name]

		print 'Cycle of cities'

		for i in range (len(dist_from_city)):	
			dist_from_city = city_name[closest_city]
			for i in range (len(visited_cities)):
				if visited_cities[i] in dist_from_city:
					dist_from_city.pop(visited_cities[i])

			sorted_dict = min(dist_from_city.items(),key=lambda x: x[1])
			closest_city = sorted_dict[0]
			visited_cities.append(closest_city)
			travel_path = travel_path + '-' + closest_city;
		print travel_path 


	except KeyError:
		print 'City does not exist'

except IOError:
	print 'Unable to read file. Check file name and run again'