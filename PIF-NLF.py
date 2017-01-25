import csv
import os
import numpy as np
import xlsxwriter

#---------------------------------------------------------------
#---------------------------------------------------------------
# This script is designed to allow MS/MS scans aquired in data dependent mode
# To be screened for specific product ions or neutral losses.
# When both product ions and neutral losses are added to the include.txt file
# Scans with both features will be searched for.
# Type the accurate mass neutral losses and product ions in the include.txt file and
# Fill in the variables below


# What is the name of the mzml text file?
mzml_file = "KAS2925-hex-DM-stepped.txt"
# What is the mass accuracy (product ion)?
ppm = 5.0
# What is the mass accuracy (neutral loss)?
mDa = 5.0
# What is the minimum intensity of the targeted product ions relative to the spectrum base peak?
min_product_peak = 45
# What is the minimum intensity of all peaks in a spectrum to be printed out (relative base peak)
min_all_peak = 5

#---------------------------------------------------------------
#---------------------------------------------------------------














# Initializing counting variables inputs
min_product_peak = min_product_peak/100.00
min_all_peak = min_all_peak/100.00
Scan_count = 0
col = 0
row = 1
N = 0
J = 0
i = 0
ppm = float(ppm)
PPM = float(ppm)
PPM = PPM/1000000.00
mDa = mDa/1000

# Reading in the data and include ions files
g = open(mzml_file, 'r')
a = g.readlines()
f = open("include.txt", "r")
include_list = f.readlines()
f.close()
g.close()


# Initializing array for productions to look for
include_product_ions = []
product_ion_include_index_start = include_list.index("#\n",0)+2
if (include_list[product_ion_include_index_start]) == "#\n":
	include_products_count = 0
		
else:
	product_ion_include_index_end = include_list.index("#\n",1)
# Number of product ions to search for
	include_products_count = product_ion_include_index_end - product_ion_include_index_start
if include_products_count != 0:
	include_list_products = include_list[product_ion_include_index_start:product_ion_include_index_end]
	J = 0
	while J < include_products_count:
		line = (include_list_products[J])
		line = line.replace("\n","")
		if line == "":
			include_products_count = include_products_count - 1
		else:
			line = float(line)
			include_product_ions.append(line)
		J = J + 1


if include_products_count == 0:
	print("No product ions will be searched for")
else:
	print("The following product ions will be looked for:")
	print(include_product_ions)
	
	
# Initializing array for neutral ions to look for		
include_neutral_loss = []
neutral_loss_include_index_start = include_list.index("#\n",1)+2
if len(include_list) > neutral_loss_include_index_start:
# Number of product ions to search for
	include_neutral = len(include_list) - neutral_loss_include_index_start
	include_list_neutral = include_list[neutral_loss_include_index_start:]
	include_list_neutral_n = len(include_list_neutral)
	
	J = 0
	while J < include_list_neutral_n :
		line = (include_list_neutral[J])
		line = line.replace("\n","")
		
		if line == "":

			include_list_neutral_n  = include_list_neutral_n  - 1
		else:

			line = float(line)
			include_neutral_loss.append(line)


		J = J + 1
#	print("The following neutral loss(es) will be searched for:")
#	print(include_list_neutral)


	

if include_neutral_loss == []:
	print("No neutral losses will be searched for")
	include_list_neutral_n = 0
else:
	print("The following neutral loss(es) will be searched for:")
	print(include_list_neutral)
		
# Opening an excel workbook to output to and writting the header titles
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
	
worksheet.write(0, col+1, "Precursor m/z")
worksheet.write(0, col+3, "Product ion m/z")	
worksheet.write(0, col+4, "Relative Intensity")	
worksheet.write(0, col+2, "Retention time")	


# Finding flags for the intensity and scan numbers 
scan_i = "        cvParam: ms level, 2\n"
mz_i = "          cvParam: m/z array, m/z\n"
intensity_i = "          cvParam: intensity array, number of detector counts\n"
# Total number of mz scans in file
MS2_scans = a.count(scan_i)

j = 0
Product_ion_count = 0
i = 0

# --------------------------------------------------
# Getting basic info about the scan
while Scan_count < MS2_scans:

# Scan location
	i = a.index(scan_i, i+1)
	ii = a.index(mz_i, i) + 1
	iv = a.index(intensity_i, i) + 1
# mz array	
	MS2_mz_string = a[ii]
	MS2_mz_end = MS2_mz_string.index("]") + 2
	MS2_mz_string = MS2_mz_string[MS2_mz_end:]
	MS2_mz_string = MS2_mz_string.replace(" ",",")
	MS2_mz = np.fromstring(MS2_mz_string, dtype=float, sep=',')
	MS2_neutral_array = MS2_mz
# Precursor ion
	iii = a.index("            isolationWindow:\n",i) + 1
	Selected_ion = a[iii]
	Selected_ion = Selected_ion.replace("              cvParam: isolation window target m/z, " , "")
	Selected_ion = Selected_ion.replace(", m/z\n" , "")
	Selected_ion = float(Selected_ion)
# Retention time	
	RT_i = a.index("          scan:\n",i) + 1
	Retention_time = a[RT_i] 
	Retention_time = Retention_time.replace("cvParam: scan start time,", "")
	Retention_time = Retention_time.replace("minute", "")
	Retention_time = Retention_time.replace(",", "")
	Retention_time = Retention_time.replace("\n", "")
	Retention_time = Retention_time.replace(" ", "")
	Retention_time = float(Retention_time)
# Neutral loss array
	MS2_neutral = [Selected_ion - x for x in MS2_neutral_array]		
	MS2_neutral = np.asarray(MS2_neutral)
	
#------------------------------------------------------------------------------------
	MS2_int_string = a[iv]
	MS2_int_end = MS2_int_string.index("]") + 2
	MS2_int_string = MS2_int_string[MS2_int_end:]
	MS2_int_string = MS2_int_string.replace(" ",",")
	MS2_int = np.fromstring(MS2_int_string, dtype=float, sep=',')	
	base_peak = max(MS2_int)
	total_mz_ions = len(MS2_mz) - 1
			
# Converting all mz values to percentages of base peak	
	I = 0	
	while I < total_mz_ions:
		MS2_int[I] = MS2_int[I]/base_peak	
		I = I + 1	

# This loop is only if a product ion is being searched for
	j = 0
	Product_ion_count = 0
	if include_products_count != 0:	
		while j < include_products_count:
			product_ion = include_product_ions[j]	
			ppm = PPM
#Determining the mass range for acceptable detection of product ions
			ppm = ppm*product_ion
			Ion_mz_low = product_ion - ppm
			Ion_mz_high = product_ion + ppm
# Counting the number of times that ion is found in the scan
			product_ion_count = ((Ion_mz_low < MS2_mz ) & (MS2_mz  < Ion_mz_high)).sum()
# Filtering based on inputted intensity requirements					
			if product_ion_count > 0:
				product_i = min(range(len(MS2_mz)), key=lambda i: abs(MS2_mz[i]-product_ion))
				product_intensity = MS2_int[product_i]
# Will count it if it is above a certain threshold			
				if product_intensity > min_product_peak:											
					Product_ion_count = Product_ion_count + product_ion_count						
# If no neutral loss filter is applied, the scan is printed out if it satisfy product ion criteria
					if include_list_neutral_n == 0:			
						I = 0
						while I < total_mz_ions:		
							MS2_intensity = MS2_int[I]
							if MS2_intensity > min_all_peak:	
								worksheet.write(row, col, N)	
								worksheet.write(row, col+1, Selected_ion)
								worksheet.write(row, col+3, MS2_mz[I])	
								worksheet.write(row, col+4, MS2_int[I])	
								worksheet.write(row, col+2, Retention_time)	
								row = row + 1
								N = N + 1

							I = I + 1
			
# If a neutral loss filter in addition to product ion filter was applied				
					else:
						neutral_count = 0
						Neutral_losses_detected = 0
						while neutral_count < include_list_neutral_n:
							neutral_ion = include_neutral_loss[neutral_count]
							Ion_mz_low = neutral_ion - mDa
							Ion_mz_high = neutral_ion + mDa								
							neutral_loss_count = ((Ion_mz_low < MS2_neutral ) & (MS2_neutral  < Ion_mz_high)).sum()
							Neutral_losses_detected = Neutral_losses_detected + neutral_loss_count						
							neutral_count = neutral_count + 1

# The targeted neutral loss was detected in the scan with the product ion of interest
		
						if Neutral_losses_detected > 0:				
							I = 0
							while I < total_mz_ions:
								MS2_intensity = MS2_int[I]
								if MS2_intensity > min_all_peak:	
									worksheet.write(row, col, N)	
									worksheet.write(row, col+1, Selected_ion)
									worksheet.write(row, col+3, MS2_mz[I])	
									worksheet.write(row, col+4, MS2_int[I])	
									worksheet.write(row, col+2, Retention_time)	
									row = row + 1
									N = N + 1
								I = I + 1					
			j = j + 1


	if include_products_count == 0:	
		neutral_count = 0
		Neutral_losses_detected = 0
		while neutral_count < include_list_neutral_n:
			neutral_ion = include_neutral_loss[neutral_count]
			Ion_mz_low = neutral_ion - mDa
			Ion_mz_high = neutral_ion + mDa								
			neutral_loss_count = ((Ion_mz_low < MS2_neutral ) & (MS2_neutral  < Ion_mz_high)).sum()
			Neutral_losses_detected = Neutral_losses_detected + neutral_loss_count						
			neutral_count = neutral_count + 1
			if Neutral_losses_detected > 0:				
				I = 0
				while I < total_mz_ions:
					MS2_intensity = MS2_int[I]
					if MS2_intensity > min_all_peak:	
						worksheet.write(row, col, N)	
						worksheet.write(row, col+1, Selected_ion)
						worksheet.write(row, col+3, MS2_mz[I])	
						worksheet.write(row, col+4, MS2_int[I])	
						worksheet.write(row, col+2, Retention_time)	
						row = row + 1
						N = N + 1
					I = I + 1	




					
	Scan_count = Scan_count + 1
workbook.close()
