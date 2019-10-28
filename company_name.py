import pandas as pd

months = {
1:"Jan",
2:"Feb",
3:"Mar",
4:"Apr",
5:"May",
6:"Jun",
7:"Jul",
8:"Aug",
9:"Sep",
10:"Oct",
11:"Nov",
12:"Dec"
}

employees = [] # List of employee names goes here

####### The first parser checks to see if the "Name" entry is an employee or a vendor name.

def first_parser(data):

	vendors, count1, count2 = [], 0, 0
	for item in data["Name"]:
		temp = str(item).replace(" ","").upper()
		if (temp in employees) or (temp == "NAN"):
			vendors.append("TEMP")
		else:
			vendors.append(str(item))
			if ":" in temp:
				count1 += 1
			if "-" in temp:
				count2 += 1
				
	for idx, vendor in enumerate(vendors):
		if (vendor != "TEMP") and (count1 < count2):
			vendor = vendor.split("-")
			vendor = vendor[0].upper()
			vendors[idx] = vendor
		elif (vendor != "TEMP"):
			vendor = vendor.split(":")
			vendor = vendor[0].upper()
			vendors[idx] = vendor
	return vendors

###### The second parser checks for "NAN" in the "Memo/Description" column, count1 and count2 determine the format of the vendor (used for splitting the vendor and description).

def second_parser(data):

	vendors, count1, count2 = [], 0, 0
	for item in data["Memo/Description"]:
		temp = str(item).replace(" ","").upper()
		if temp == "NAN":
			vendors.append("TEMP")
		else:
			vendors.append(str(item))
			if ":" in temp:
				count1 += 1
			if "-" in temp:
				count2 += 1
				
	for idx, vendor in enumerate(vendors):
		if (vendor != "TEMP") and (count1 < count2):
			vendor = vendor.split("-")
			vendor = vendor[0].upper()
			vendors[idx] = vendor
		elif (vendor != "TEMP"):
			vendor = vendor.split(":")
			vendor = vendor[0].upper()
			vendors[idx] = vendor
	return vendors
		
###### The third parser takes compares the results of the two previous parsers and returns a single list of vendor names

def third_parser(lst1,lst2):
	
	vendors = []
	for idx, item in enumerate(lst1):
		if item != "TEMP":
			vendors.append(item)
		elif lst2[idx] != "TEMP":
			vendors.append(lst2[idx])
		else:
			vendors.append("UNKNOWN")
    return vendors
