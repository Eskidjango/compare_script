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

# # These cannot be empty
# OLD_VERSION = 'to_comp_photosavings/NEW VERSION.csv'
# NEW_VERSION = 'to_comp_photosavings/OLD VERSION.csv'

# # Can be empty (in this case missing data will not have a reason description in output file)
# DEBUG_REPORT_REIMBURSEMENTS = 'to_comp_photosavings/REIMBURSEMENTS.csv'
# DEBUG_REPORT_RETURNS_TO_FBA = 'to_comp_photosavings/RETURNS_TO_FBA.csv'
# DEBUG_REPORT_DATE_RANGE = 'to_comp_photosavings/DATE_RANGE_ORDER.csv'



class OrderDiscrepancyComparisonScript:
	"""
	Compare order discrepancy reports and put the result into an output file 
	named as 'Order_Discrepancy_compare_result.xlsx'
	This class gets data from old_version and new_version, compare these versions and 
	compute, which items are missing, which are not. Then it's looking for a reason of missing in debug report 
	and put its description into the output file.
	

	:param old_version: A csv file of old version of the report.
	:param new_version: A csv file of new version of the report. 	
	:param reimbursements: A csv file of 'Reimbursements' sheet of the debug report
	:param returns_to_fba: A csv file of 'ReturnsToFBA' sheet of the debug report
	:param date_range: A csv file of 'DateRangeCSV' sheet of the debug report
	"""

	reasons = {
		'unknown_reason': 'unknown reason',
		'not_in_debug': 'not in debug report',
		'refund_without_order': 'In Date Range this order has only a refund record without an order one',
		'refunds_qty_greater_than_orders': 'Refunds quantity cannot be greater than orders so variance <= 0 here',
		'should_be_in_report': 'Should be in report because variance > 0 here'
		}

	# Statuses
	MISSING = 'MISSING'
	FINE = 'FINE'

	def __init__(self, old_version, new_version, reimbursements=None, returns_to_fba=None, date_range=None):
		self.old_version = old_version
		self.new_version = new_version

		self.reimbursements = reimbursements
		self.returns_to_fba = returns_to_fba
		self.date_range = date_range

		# Dict to store the data from old and new versions
		# order_id is a KEY and a dict of sku, order qty, refund qty, returned qty, 
		# status, missing in, reason is a VALUE
		self.reports_data = {} 

		# Lists of tuples in format [(order_id, sku), ]
		# For example: [('106-5388488-2997800, 3202NVYS_FBA'), ('111-1452330-6698653, E104BTTBL_FBA')]
		# They are needed to find missing data		
		self.old_data_order_skus = []
		self.new_data_order_skus = []



	def _read_reports_data_from_files(self):
		"""
		Read data from files and put each sku and order_id to lists
		in format 'sku#order_id' to find the missing data then
		"""

		with open(self.old_version, 'rb') as csvfile:
			old_data = csv.reader(csvfile, delimiter=',')
			next(old_data) # not to read column titles	

			for row in old_data:
				order_id = row[1]
				sku = row[2] 
				# order_id and sku might be empty
				if (order_id and sku):
					order_sku = (order_id, sku)
					self.old_data_order_skus.append(order_sku)
		
		with open(self.new_version, 'rb') as csvfile:
			new_data = csv.reader(csvfile, delimiter=',')
			next(new_data) # not to read column titles

			for row in new_data:
				order_id = row[1] 
				sku = row[2] 			
				# order_id and sku might be empty
				if (order_id and sku):
					order_sku = (order_id, sku)
					self.new_data_order_skus.append(order_sku)
		


	def _find_missing_data(self):
		"""

		Find missing data and put it to 'self.reports_data' dict with
		status 'MISSING'. Both versions data get status 'FINE'

		"""

		for order_sku in self.old_data_order_skus:			
			self.reports_data[order_sku] = {}			
			self.reports_data[order_sku]['status'] = self.FINE			
			if order_sku not in self.new_data_order_skus:				
				self.reports_data[order_sku]['status'] = self.MISSING
				self.reports_data[order_sku]['missing_in'] = 'new version (path: {})'.format(self.new_version) # In which version it's missing
				self.reports_data[order_sku]['reason'] = self.reasons['unknown_reason']				
					
		for order_sku in self.new_data_order_skus:				
			self.reports_data[order_sku] = {}			
			self.reports_data[order_sku]['status'] = self.FINE
			if order_sku not in self.old_data_order_skus:				
				self.reports_data[order_sku]['status'] = self.MISSING
				self.reports_data[order_sku]['missing_in'] = 'old version (path: {})'.format(self.old_version) # In which version it's missing
				self.reports_data[order_sku]['reason'] = self.reasons['unknown_reason']


	def _get_reports_data_info_from_debug(self):
		'''

		Getting through all the sheets of debug report (Reimbursements, ReturnsToFBA, DateRangeCSV) and
		taking info for all the items we have in 'self.reports_data' and 
		put it to the 'self.reports_data' dict 
		Info means these characteristics:
			1. Order quantity (from date range)
			2. Refund quantity (from date_range)
			3. Reimbursed quantity (from reimbursements)
			4. Returned quantity (from returns to fba)
		All of those are needed to discover the reasons of missing
		'''
		print 'Getting info from debug...'
		if not self.date_range or not self.reimbursements or not self.returns_to_fba:			
			raise EOFError, 'Some neccessary debug files weren\'t included.'\
			' Make sure you have included each of these: Reimbursements, ReturnsToFBA, DateRangeCSV from debug report'
							
		with open(self.reimbursements, 'rb') as csvfile:
			reimbursements_data = csv.reader(csvfile, delimiter=',')	
			next(reimbursements_data) # Not to read column titles

			for row in reimbursements_data:
				reimbursed_qty = int(row[15] or 0) # Reimbursed qty may be an empty string ''
				order_id, sku = row[3], row[5]
				order_sku_key = (order_id, sku)
				if order_sku_key in self.reports_data.keys():
					data = self.reports_data[order_sku_key]
					data['in_debug'] = True
					data['reimbursed_qty'] = data.get('reimbursed_qty', 0) + reimbursed_qty										
											
		with open(self.returns_to_fba, 'rb') as csvfile:
			returns_to_fba_data = csv.reader(csvfile, delimiter=',')	
			next(returns_to_fba_data) 

			for row in returns_to_fba_data:
				returned_qty = int(row[6])
				order_id, sku = row[1], row[2]
				order_sku_key = (order_id, sku)
				if order_sku_key in self.reports_data.keys():
					data = self.reports_data[order_sku_key]
					data['in_debug'] = True
					data['returned_qty'] = data.get('returned_qty', 0) + returned_qty											
					
		with open(self.date_range, 'rb') as csvfile:
			date_range_data = csv.reader(csvfile, delimiter=',')	
			next(date_range_data)

			for row in date_range_data:	
				order_condition = row[2] # 'Refund' or 'Order'
				quantity = int(row[6] or 0) # Refund or order's quantity may be an empty string ''			
				order_id, sku = row[3], row[4]
				order_sku_key = (order_id, sku)
				if order_sku_key in self.reports_data.keys():	
					data = self.reports_data[order_sku_key]	
					data['in_debug'] = True			
					if order_condition == 'Order':
						data['order_qty'] = data.get('order_qty', 0) + quantity	
					elif order_condition == 'Refund':
						data['refund_qty'] = data.get('refund_qty', 0) + quantity
			

	def _check_not_in_debug_reason(self):
		print 'Check not in debug reason'
		for order_sku, data in self.reports_data.iteritems():
			if not data.get('in_debug'):
				data['reason'] = self.reasons['not_in_debug']


	def _check_refund_without_order_reason(self):
		print 'Check refund without order reason'
		for order_sku, data in self.reports_data.iteritems():
			if data['status'] == self.MISSING and data.get('in_debug'):
				order_qty = data.get('order_qty', 0)
				refund_qty = data.get('refund_qty', 0)
				if order_qty == 0 and refund_qty > 0:
					data['reason'] = self.reasons['refund_without_order']


	def _check_refunds_qty_greater_than_orders_reason(self):
		print 'Check refunds qty greater than orders reason'
		for order_sku, data in self.reports_data.iteritems():
			if data['status'] == self.MISSING and data.get('in_debug'):
				order_qty = data.get('order_qty', 0)
				refund_qty = data.get('refund_qty', 0)
				if order_qty != 0 and order_qty < refund_qty:
					# In new logic refund qty cannot be greater than orders
					# if so, we should make refund qty the same as order qty
					refund_qty = order_qty 
					variance = refund_qty - data.get('reimbursed_qty', 0) - data.get('returned_qty', 0)
					if variance <= 0:
						data['reason'] = self.reasons['refunds_qty_greater_than_orders']


	def _check_should_be_in_report_reason(self):		
		print 'Check should be in report reason'
		for order_sku, data in self.reports_data.iteritems():
			if data['status'] == self.MISSING and data.get('in_debug'):
				variance = data.get('refund_qty', 0) - data.get('reimbursed_qty', 0) - data.get('returned_qty', 0)
				if variance > 0 and data.get('order_qty'):
					data['reason'] = self.reasons['should_be_in_report']		


	def _make_output_file(self):
		output_file = Workbook()
		output_sheet = output_file.active
		output_sheet.title = 'Auto Output'

		output_sheet.column_dimensions['A'].width = 21
		output_sheet.column_dimensions['B'].width = 20
		output_sheet.column_dimensions['C'].width = 8
		output_sheet.column_dimensions['D'].width = 8
		output_sheet.column_dimensions['E'].width = 8
		output_sheet.column_dimensions['F'].width = 8	
		output_sheet.column_dimensions['G'].width = 10	
		output_sheet.column_dimensions['H'].width = 15
		output_sheet.column_dimensions['I'].width = 40
		
		output_sheet.append(
			['Order_id', 'SKU', 'Order Qty', 'Refund Qty', 'Returned Qty', 'Reimbursed Qty', 'Status', 'Missing in', 'Reason']
			)

		result_data = []
		for order_sku, data in self.reports_data.iteritems():
			order_id, sku = order_sku
			reports_data_row = [
				order_id, 
				sku,				
				str(data.get('order_qty', ' ')), 
				str(data.get('refund_qty', ' ')),
				str(data.get('returned_qty', ' ')), 
				str(data.get('reimbursed_qty', ' ')),
				data['status'],
				data.get('missing_in'), 				
				data.get('reason')
			]
			result_data.append(reports_data_row)

		sorted_result_data_by_order_qty = sorted(result_data, key=lambda i: i[6], reverse=True)

		for item in sorted_result_data_by_order_qty:
			output_sheet.append(item)

		output_file_name = "Order_Discrepancy_comparison_result_{date}.xlsx".format(date=datetime.now())	

		output_file.save(output_file_name)

		print 'Compare files has just been finished! Check the \'{file_name}\' to see the result'.format(file_name=output_file_name)



	def run_script(self):
		self._read_reports_data_from_files()
		self._find_missing_data()		
		self._get_reports_data_info_from_debug()
		self._check_not_in_debug_reason()
		self._check_refund_without_order_reason()
		self._check_refunds_qty_greater_than_orders_reason()
		self._check_should_be_in_report_reason()
		self._make_output_file()





if __name__ == '__main__':
	Jacob_data = OrderDiscrepancyComparisonScript(
		old_version=OLD_VERSION,
    	new_version=NEW_VERSION,
	  	reimbursements=DEBUG_REPORT_REIMBURSEMENTS,
	  	returns_to_fba=DEBUG_REPORT_RETURNS_TO_FBA,
	  	date_range=DEBUG_REPORT_DATE_RANGE,
	  )
	Jacob_data.run_script()

