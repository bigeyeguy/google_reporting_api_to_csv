"""Hello Analytics Reporting API V4.
https://developers.google.com/analytics/devguides/reporting/core/v4/quickstart/installed-py
https://developers.google.com/analytics/devguides/reporting/core/v4/basics#metrics"""

from google2pandas import *
import xlsxwriter
import pandas as pd

## User Management - ADD Read & Analyse permissions for the VIEW for ______@_____.iam.gserviceaccount.com

VIEW_ID = 'ENTER_ID' 
CLIENT_NAME = 'CLIENT_NAME'
DATE = [{ "endDate": "0000-00-00", "startDate": "0000-00-00" }]
MONTH = 'ENTER_MONTH'

# All channels
query1 = {
    'reportRequests': [{
        'viewId' : VIEW_ID,
        
        'dateRanges': DATE,
            
        'dimensions' : [{"name":"ga:channelGrouping"}],
            
        'metrics'   : [
            {"expression": "ga:sessions"},
            {'expression': 'ga:newUsers'},
            {'expression': 'ga:transactionRevenue'},
			{"expression": "ga:transactions"},
			{"expression": "ga:bounceRate"}],
    }]
}

# organic landing pages (no metricFilterClauses)
query2 = {
    "reportRequests": [{
        "dateRanges": DATE,
        
		"dimensions":[
          {"name":"ga:channelGrouping"},
		  {"name":"ga:landingPagePath"}],          

        "dimensionFilterClauses": [{
            "filters": [{
                "dimension_name": "ga:channelGrouping",
                "operator": "EXACT",
                "expressions": ["Organic Search"]
            }]
        }],
		
		"metrics": [
            {"expression": "ga:sessions"},
			{"expression": "ga:percentNewSessions"},
            {'expression': 'ga:newUsers'},
			{"expression": "ga:bounceRate"},
			{"expression": "ga:pageviewsPerSession"},
			{"expression": "ga:sessionDuration"},
			{"expression": "ga:organicSearches"},
            {'expression': 'ga:transactionRevenue'},
			{"expression": "ga:transactions"}],
		
		
        "viewId": VIEW_ID,
        
    }]
}

# device
query3 = {
        "reportRequests": [{
      "dateRanges": DATE,
      "metrics": [
          {"expression": "ga:sessions"},
          {'expression': 'ga:transactionRevenue'},
          {'expression': 'ga:revenuePerTransaction'},
      ],
      "viewId": VIEW_ID,
      "dimensions":[
        {"name":"ga:deviceCategory"},
	  ],
  }]
}


# site search
query4 = {
    "reportRequests": [{
        "dateRanges": DATE,
        
		"dimensions":[
		  {"name":"ga:searchKeyword"}],          

		
		"metrics": [
            {"expression": "ga:searchUniques"},
			{"expression": "ga:avgSearchResultViews"},
            {'expression': 'ga:percentSearchRefinements'},
			{"expression": "ga:searchExits"}
			],
		
		"metricFilterClauses": [{
            "filters": [{
                "metricName": "ga:searchUniques",
                "operator": "GREATER_THAN",
                "comparisonValue": "3"
            }]
        }],
			
        "viewId": VIEW_ID,
        
    }]
}

# Top Products
query5 = {
    "reportRequests": [{
        "dateRanges": DATE,
        
		"dimensions":[
          {"name":"ga:productName"}],          

		
		"metrics": [
            {"expression": "ga:itemRevenue"},
			{"expression": "ga:itemQuantity"},
            {'expression': 'ga:uniquePurchases'},
			{"expression": "ga:revenuePerItem"},
			{"expression": "ga:itemsPerPurchase"}],
		
		# metric filter returns landing pages w/ >20 sessions 
		# orderBys doesn't not work 

		
		
        "viewId": VIEW_ID,
        
    }]
}
# device - organic
query6 = {
        "reportRequests": [{
      "dateRanges": DATE,
      "metrics": [
          {"expression": "ga:sessions"},
          {'expression': 'ga:transactionRevenue'},
          {'expression': 'ga:revenuePerTransaction'},
      ],
      "viewId": VIEW_ID,
      "dimensions":[
        {"name":"ga:deviceCategory"},
        {"name":"ga:channelGrouping"}
	  ],
	  
	  "dimensionFilterClauses": [{
            "filters": [{
                "dimension_name": "ga:channelGrouping",
                "operator": "EXACT",
                "expressions": ["Organic Search"]
            }]
        }],
  }]
}
    
# Assume we have placed our client_secrets_v4.json file in the current working directory.

conn = GoogleAnalyticsQueryV4(secrets='_______________.json')
df1 = conn.execute_query(query1)
df2 = conn.execute_query(query2)
df3 = conn.execute_query(query3)
df4 = conn.execute_query(query4)
df5 = conn.execute_query(query5)
df6 = conn.execute_query(query6)


df1 = df1[['channelGrouping', 'sessions', 'newUsers', 'transactionRevenue', 'bounceRate', 'transactions']]
df2 = df2[['landingPagePath', 'channelGrouping', 'sessions', 'newUsers', 'bounceRate', 'organicSearches', 'pageviewsPerSession', 'percentNewSessions', 'sessionDuration', 'transactions', 'transactionRevenue']]
df4 = df4[['searchKeyword', 'searchUniques', 'avgSearchResultViews', 'searchExits', 'percentSearchRefinements']]
df5 = df5[['productName', 'itemRevenue', 'uniquePurchases', 'itemQuantity', 'revenuePerItem', 'itemsPerPurchase' ]]

# re-order columns ^^
# but I new need to reorder the column formats and array formulas

writer = pd.ExcelWriter('REPORT_NAME_%s_%s.xlsx' % (CLIENT_NAME, MONTH), engine = 'xlsxwriter', options={'strings_to_numbers':True})
df1.to_excel(writer, sheet_name = 'All Channels', index = False)
df2.to_excel(writer, sheet_name = 'Top Landing Pages', index = False)
df3.to_excel(writer, sheet_name = 'Device - All', index = False)
df4.to_excel(writer, sheet_name = 'Internal Search', index = False)
df5.to_excel(writer, sheet_name = 'Top Products', index = False)
df6.to_excel(writer, sheet_name = 'Device - Organic', index = False)


workbook = writer.book
worksheet1 = writer.sheets['All Channels']
worksheet2 = writer.sheets['Top Landing Pages']
worksheet3 = writer.sheets['Device - All']
worksheet4 = writer.sheets['Internal Search']
worksheet5 = writer.sheets['Top Products']
worksheet6 = writer.sheets['Device - Organic']

# https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html
# https://xlsxwriter.readthedocs.io/format.html

format1 = workbook.add_format({'num_format': '0.00'})
format2 = workbook.add_format({'num_format': '#,##0'})
format3 = workbook.add_format({'num_format': '#,##0.00'})
bold = workbook.add_format({'bold':True, 'bottom':True})


worksheet1.set_column('B:C', 20, format2)
worksheet1.set_column('D:D', 20, format3)
worksheet1.set_column('E:E', 20, format1)
worksheet1.set_column('F:H', 20)
worksheet1.write('G1', 'Conversion Rate', bold)
worksheet1.write('H1', 'AOV', bold)
worksheet1.write_array_formula('G2:G15', '{=f2:f15/b2:b15*100}', format1)
worksheet1.write_array_formula('H2:H15', '{=d2:d15/f2:f15}', format1) 
worksheet1.set_column('G:H', 20, format1)

worksheet1.write('A20', 'Total')
worksheet1.write_array_formula('B20', '=SUM(B2:B15)', format2)
worksheet1.write_array_formula('C20', '=SUM(C2:C15)', format2)
worksheet1.write_array_formula('D20', '=SUM(D2:D15)', format3)
worksheet1.write_array_formula('F20', '=SUM(F2:F15)', format2)
worksheet1.write_formula('G20', '=SUM(F20/B20*100)', format1)
worksheet1.write('E20', 'Get from GA')
worksheet1.write('H20', '=sum(D20/F20)', format3)

worksheet2.set_column('C:D', 12, format2)
worksheet2.set_column('E:E', 12, format1)
worksheet2.set_column('F:F', 25, format2)
worksheet2.set_column('G:H', 25, format1)
worksheet2.set_column('K:K', 25, format3)
worksheet2.set_column('A:A', 35)
worksheet2.set_column('I:J', 25)

worksheet3.set_column('B:B', 25, format1)
worksheet3.set_column('C:C', 25, format2)
worksheet3.set_column('D:D', 25, format3)
worksheet3.set_column('A:A', 25)

worksheet4.set_column('C:C', 25, format1)
worksheet4.set_column('E:E', 25, format1)
worksheet4.set_column('A:A', 35)
worksheet4.set_column('B:E', 25)


worksheet5.set_column('B:B', 25, format3)
worksheet5.set_column('E:E', 25, format3)
worksheet5.set_column('F:F', 25, format1)
worksheet5.set_column('A:A', 35)
worksheet5.set_column('B:F', 25)

writer.save()



