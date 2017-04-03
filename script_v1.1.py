from datetime import datetime
import csv
from openpyxl import Workbook

"""

HOW TO USE IT:

1. Put file paths to the variables below
2. Run the script
3. The result file will appear in the same directory where this script is

"""

# These cannot be empty
OLD_VERSION = 'to_comp_Jacob/51.csv'
NEW_VERSION = 'to_comp_Jacob/26.csv'

# Can be empty (in this case missing data will not have a reason description in output file)
DEBUG_REPORT_REIMBURSEMENTS = 'to_comp_Jacob/reimbursements.csv'
DEBUG_REPORT_RETURNS_TO_FBA = 'to_comp_Jacob/returns_to_fba.csv'
DEBUG_REPORT_DATE_RANGE = 'to_comp_Jacob/date_range.csv'



def compare_order_discrepancy_report_versions(old_version, 
											  new_version, 											   
											  reimbursements=None, 
											  returns_to_fba=None,
											  date_range=None
											  ):
	"""
	Compare order discrepancy reports and put the result into an output file 
	named as 'Order_Discrepancy_compare_result.xlsx'
	This function get data from old_version, get data from new_version, compare these versions and 
	compute, which items are missing, which are not. Then it's looking for a reason of missing in debug report 
	and put its description into the output file.
	

	:param old_version: A csv file of old version of the report.
	:param new_version: A csv file of new version of the report. 	
	:param reimbursements: A csv file of 'Reimbursements' sheet of the debug report
	:param returns_to_fba: A csv file of 'ReturnsToFBA' sheet of the debug report
	:param date_range: A csv file of 'DateRangeCSV' sheet of the debug report
	"""

	# Lists's items are rows that will be put into output file
	# Items have the following format: ['Order_id', 'SKU', 'Not in', 'reason']
	# For example: ['106-5388488-2997800', '3202NVYS_FBA', 'new version (path: to_compare/ORDER_NEW_VERS.csv)',
	# 'unknown_reason']
	missing_records = []
	both_versions_records = [] # not missing records

	# Lists of order keys in format 'order_id#sku'
	# They are needed to find out missing data
	# For example: ['106-5388488-2997800#3202NVYS_FBA', '111-1452330-6698653#E104BTTBL_FBA']	
	old_data_order_keys = []
	new_data_order_keys = []

	reasons = {
		'unknown_reason': 'unknown reason',
		'not_in_debug_report': 'An order with that sku hasn\'t been found in debug report',
		'refund_without_order': 'In Date Range this order has only a refund record without an order one',
		'refunds_qty_greater_than_orders': 'Refunds are more than orders, but we don\'t count all of them because'
										' the quantity of orders cannot be greater than qunatity of refunds by this logic'
										' (see \'Not in\' column). So variance here is <=0'		
	}

	# Make an output file to put the data into it
	output_file = Workbook()
	output_sheet = output_file.active
	output_sheet.title = 'Auto Output'

	# column titles
	output_sheet.append(
		['Order_id', 'SKU', 'Not in', 'Reason']
		)

	output_sheet.column_dimensions['A'].width = 21
	output_sheet.column_dimensions['B'].width = 20
	output_sheet.column_dimensions['C'].width = 15
	output_sheet.column_dimensions['D'].width = 25

	# Getting all order keys from an old version of the report and put them into list
	with open(old_version, 'rb') as csvfile:
		old_data = csv.reader(csvfile, delimiter=',')
		next(old_data) # not to read column titles
		
		for row in old_data:
			order_id = row[1]
			sku = row[2] 
			# order_id and sku might be empty in some rows, so we need to check it
			if (order_id and sku):
				order_key = '{}#{}'.format(order_id, sku)
				old_data_order_keys.append(order_key)

	# Getting all order keys from a new version of the report and put them into list
	with open(new_version, 'rb') as csvfile:
		new_data = csv.reader(csvfile, delimiter=',')
		next(new_data) # not to read column titles

		for row in new_data:
			order_id = row[1] 
			sku = row[2] 			
			# order_id and sku might be empty in some rows, so we need to check it
			if (order_id and sku):
				order_key = '{}#{}'.format(order_id, sku)
				new_data_order_keys.append(order_key)

	# Computing which elements are missing, which are not	
	for item in old_data_order_keys:
		# Split order key to put its values into separated cells in the output file	
		order_id = item.split('#')[0] 
		sku = item.split('#')[1] 				
		if item not in new_data_order_keys:			
			not_in = 'new version (path: {})'.format(new_version) # In which version it's missing	
			# Put missing elements into 'missing_records' list	
			missing_data = [order_id, sku, not_in]					
			missing_records.append(missing_data)
		else:
			# Put not missing elements into 'both_versions_records' list
			fine_data = [order_id, sku, '-', '-']
			# Check if it was added already
			if fine_data not in both_versions_records:
				both_versions_records.append(fine_data)	

	# Computing which elements are missing, which are not			
	for item in new_data_order_keys:
		# Split order key to put its values into separated cells in the output file	
		order_id = item.split('#')[0]
		sku = item.split('#')[1]	
		if item not in old_data_order_keys:
			not_in = 'old version (path: {})'.format(old_version) # In which version it's missing	
			# Put missing elements into 'missing_records' list	
			missing_data = [order_id, sku, not_in]					
			missing_records.append(missing_data)
		else:			
			fine_data = [order_id, sku, '-', '-']
			# Check if it was added already
			if fine_data not in both_versions_records:
				both_versions_records.append(fine_data)	



	checked_data = _check_if_items_are_in_debug(
		missing_records, 
		reimbursements=reimbursements, 
		returns_to_fba=returns_to_fba, 
		date_range=date_range
		)
	print checked_data

	# Check 'refund_without_order' and 'refunds_qty_greater_than_orders' reasons
	if date_range:
		with open(date_range, 'rb') as csvfile:
			debug_data = csv.reader(csvfile, delimiter=',')			
			# Take each item from 'missing_records' list and find all records with the same order_id and sku in 'Date Range'
			# Then count order_condition (3rd column) of the item. Its value can be 'Refund' or 'Order'
			# If count of order records is 0 and there are refund records, then put the reason 'refund_without_order'
			for item in missing_records:
				if item not in checked_data:					
					order_id = item[0] 
					sku = item[1] 			
					refund_records_quantity = 0 # To count the number of 'Refund' ones  in 'DateRangeCSV' sheet
					order_records_quantity = 0 # To count the number of 'Order' ones in 'DateRangeCSV' sheet
					# iterate through 'DateRangeCSV' sheet of the debug report to see order_condition of the item			
					for row in debug_data:	
						order_condition = row[2] # 3rd column in 'DataRangeCSV'. Cell value can be 'Order' or 'Refund'			
						if order_id in row and sku in row and order_condition == 'Refund':
							refund_records_quantity += 1
						elif order_id in row and sku in row and order_condition == 'Order':
							order_records_quantity += 1
					if order_records_quantity == 0 and refund_records_quantity > 0:
						item.append(reasons['refund_without_order'])	
						checked_data.append(item)
					elif order_records_quantity > refund_records_quantity:
						item.append(reasons['refunds_qty_greater_than_orders'])
						checked_data.append(item)
					else:
						item.append(reasons['unknown_reason'])	
						checked_data.append(item)
					csvfile.seek(0)	# Back to the begginning of the file


						
	# Write data into the output file
	for item in checked_data:
		output_sheet.append(item)

	for item in both_versions_records:
		output_sheet.append(item)
	

	output_file_name = "Order_Discrepancy_compare_result_{date}.xlsx".format(date=datetime.now())	

	output_file.save(output_file_name)

	print 'Compare files has just been finished! Check the \'{file_name}\' to see the result'.format(file_name=output_file_name)



def _check_if_items_are_in_debug(missed_data, reimbursements, returns_to_fba, date_range):
	# Expect if date_range doesn't have certain order id and sku then returns_to_fba doesn't too.	
	checked_data = []
		
	for item in missed_data:
		order_id = item[0]
		sku = item[1]

		item_in_debug = False
		with open(reimbursements, 'rb') as csvfile:
				reimbursements_data = csv.reader(csvfile, delimiter=',')	
				next(reimbursements_data)

				for row in reimbursements_data:
					if order_id in row and sku in row:
						# print 'reimbursements', order_id, sku
						item_in_debug = True

		with open(returns_to_fba, 'rb') as csvfile:
				returns_to_fba_data = csv.reader(csvfile, delimiter=',')	
				next(returns_to_fba_data)

				for row in returns_to_fba_data:
					if order_id in row and sku in row:
						# print 'returns_to_fba', order_id, sku
						item_in_debug = True

		with open(date_range, 'rb') as csvfile:
				debug_data = csv.reader(csvfile, delimiter=',')	
				next(debug_data)

				for row in debug_data:
					if order_id in row and sku in row:
						# print 'date_range', order_id, sku, row[2]
						item_in_debug = True

		if not item_in_debug:
			item.append('Item is not in debug')		
			checked_data.append(item)

	return checked_data



if __name__ == '__main__':
	compare_order_discrepancy_report_versions(old_version=OLD_VERSION,
											  new_version=NEW_VERSION,
											  reimbursements=DEBUG_REPORT_REIMBURSEMENTS,
											  returns_to_fba=DEBUG_REPORT_RETURNS_TO_FBA,
											  date_range=DEBUG_REPORT_DATE_RANGE,
											  )
	
