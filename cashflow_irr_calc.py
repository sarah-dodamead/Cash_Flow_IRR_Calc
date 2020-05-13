import xlrd
import pandas as pd
import numpy as np

#loading exvel sheet
loc = ("/Users/Nomad/Downloads/Cash Flow.xlsx") #path in my local mac to cash flow sheet
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#variables with inputs from xls sheet
solar_PV_CAPEX = sheet.cell_value(4,2)
solar_PV_year = sheet.cell_value(4,3)
solar_PV_revenue_yr1 = sheet.cell_value(4,6)
solar_PV_revenue_yr1_annual_escalation = sheet.cell_value(4,7)
solar_PV_OPEX_yr1 = sheet.cell_value(4,11)
solar_PV_OPEX_yr1_annual_escalation = sheet.cell_value(4,12)
federal_income_tax = sheet.cell_value(4,15)
BESS_CAPEX_1 = sheet.cell_value(6,2)
BESS_CAPEX_1_yr = sheet.cell_value(6,3)
BESS_CAPEX_2 = sheet.cell_value(7,2)
BESS_CAPEX_2_yr = sheet.cell_value(7,3)
BESS_revenue_yr1 = sheet.cell_value(5,6)
BESS_revenue_yr1_annual_escalation = sheet.cell_value(5,7)
BESS_OPEX_yr1 = sheet.cell_value(5,11)
BESS_OPEX_yr1_annual_escalation = sheet.cell_value(5,12)

# Data frame of Input values
data = [{'CAPEX':solar_PV_CAPEX, 'YEAR': solar_PV_year, 'Year1 rev': solar_PV_revenue_yr1 , 'AE CAPEX':solar_PV_revenue_yr1_annual_escalation, 'YEAR1 OPEX':solar_PV_OPEX_yr1, 'AE OPEX':solar_PV_OPEX_yr1_annual_escalation, 'FIT':federal_income_tax}, {'Year1 rev': BESS_revenue_yr1 , 'AE CAPEX':BESS_revenue_yr1_annual_escalation, 'YEAR1 OPEX':BESS_OPEX_yr1, 'AE OPEX':BESS_OPEX_yr1_annual_escalation}, {'CAPEX':BESS_CAPEX_1, 'YEAR': BESS_CAPEX_1_yr}, {'CAPEX':BESS_CAPEX_2,'YEAR':BESS_CAPEX_2_yr} ]
df = pd.DataFrame(data, columns = ['CAPEX','YEAR', 'Year1 rev', 'AE CAPEX', 'YEAR1 OPEX', 'AE OPEX', 'FIT' ], index = ['solar','bess', 'bess', 'bess'])
print(df)

years = np.arange(21) #array of years 0-20 to evaluate time series evolution of cashflow, years will be used in the following for loops

#initializing all of the arrays that will be used to build the df
CAPEXpv_array = []
CAPEXbess_array = []
total_capex = []
REVpv_array = []
REVbess_array = []
totalrev_array = []
opexpv_array = []
opexbess_array = []
totalopex_array = []
solar_income_array = []
bess_income_array = []
solar_tax_total = []
bess_tax_total = []
solar_income_after_tax_array = []
bess_income_after_tax_array = []
total_income_after_tax_array = []
NCF_solar_array = []
NCF_bess_array = []
PCF_array = []

#CAPEX calculations
#CAPEX PV

for i in years: #for loop to itterate through every year 
	if i == solar_PV_year : # year 0 value comes from inputs
		PV_capex = solar_PV_CAPEX *(-1)
	else :
		PV_capex = 0
	CAPEXpv_array.append(PV_capex)

#CAPEX BESS

for i in years:
	if i == BESS_CAPEX_1_yr :
		bess_capex = BESS_CAPEX_1*(-1) # year 0 value comes from inputs
	elif i == BESS_CAPEX_2_yr :
		bess_capex = BESS_CAPEX_2*(-1)# year 10 value comes from inputs
	else :
		bess_capex = 0 #CAPEX for every year but 0 and 10
	CAPEXbess_array.append(bess_capex)
	total_capex_i = CAPEXbess_array[i] + CAPEXpv_array[i] #CAPEX of BESS+PV
	total_capex.append(total_capex_i)

#Revenue calculations

for i in years:
	if i == 0 :
		bess_rev = np.nan #year 0 has no reveunes for bess
		PV_rev = np.nan #year 0 has no reveunes for pv
	elif i == 1 : # year one revenue is an input value calculated here
		bess_rev = BESS_revenue_yr1
		PV_rev = solar_PV_revenue_yr1
	else : #year 2-20 is calculated here
		PV_rev = round((solar_PV_revenue_yr1*(1+solar_PV_revenue_yr1_annual_escalation)**(years[i]-1)),2) #used round function so the calculated values would round to 2 decimal places to mimic money, can be changed to 0 to more closly match excel format
		bess_rev = round((BESS_revenue_yr1*(1+BESS_revenue_yr1_annual_escalation)**(years[i]-1)),2) #used round function so the calculated values would round to 2 decimal places to mimic money, can be changed to 0 to more closly match excel format
	REVpv_array.append(PV_rev) #REVENUE array for pv being created
	REVbess_array.append(bess_rev) #REVENUE array for bess being created
	total_rev = REVbess_array[i] + REVpv_array[i]
	totalrev_array.append(total_rev) # total revenue (bess+pv)

#OPEX calculations

for i in years:
	if i == 0 : #year 0 has no opex
		PV_opex = np.nan
		bess_opex = np.nan
	elif i == 1 :
		PV_opex = solar_PV_OPEX_yr1*(-1) #opex pv year 1
		bess_opex = BESS_OPEX_yr1*(-1) #opex bess year 1
	else :
		bess_opex = round((BESS_OPEX_yr1*(-1)*(1+BESS_OPEX_yr1_annual_escalation)**(years[i]-1)),2) #used round function so the calculated values would round to 2 decimal places to mimic money, can be changed to 0 to more closly match excel format
		PV_opex = round((solar_PV_OPEX_yr1*(-1)*(1+solar_PV_OPEX_yr1_annual_escalation)**(years[i]-1)),2) #pv opex calculation
	opexpv_array.append(PV_opex)
	opexbess_array.append(bess_opex)

#total opex, sum of bess+solar
for i in years: #year 0 has no opex
	if i == 0 :
		total_opex = np.nan
	else :
		total_opex = round((opexbess_array[i] + opexpv_array[i]),2)
	totalopex_array.append(total_opex)
	

#INCOME TAX CALCULATION
#solar and bess taxable income calcualtion

for i in years:
	if i == 0: #year 0 has no opex
		solar_income = np.nan
		bess_income = np.nan
	else:
		solar_income = round((opexpv_array[i] + REVpv_array[i]),2) #used round function so the calculated values would round to 2 decimal places to mimic money, can be changed to 0 to more closly match excel format
		bess_income = round((opexbess_array[i] + REVbess_array[i]),2)
	bess_income_array.append(bess_income) #BESS taxable income
	solar_income_array.append(solar_income) #PV taxable income


#Solar pv and bess tax

for i in years:
	if i == 0: #year 0 has no opex
		solar_tax = np.nan
		bess_tax = np.nan
	else :
		solar_tax = round((federal_income_tax * solar_income_array[i] * (-1)),2) #used round function so the calculated values would round to 2 decimal places to mimic money, can be changed to 0 to more closly match excel format
		bess_tax = round((federal_income_tax * bess_income_array[i] * (-1)),2)
	solar_tax_total.append(solar_tax) #pv tax on revenue
	bess_tax_total.append(bess_tax) #bess tax on revenue

#income after tax


for i in years:
	if i == 0: #year 0 has no opex
		s_income_after_tax = np.nan
		b_income_after_tax = np.nan
	else:
		s_income_after_tax = solar_tax_total[i] + solar_income_array[i] #solar income after tax
		b_income_after_tax = bess_tax_total[i] + bess_income_array[i] #bess income after tax
	solar_income_after_tax_array.append(s_income_after_tax) #generating array for after tax income
	bess_income_after_tax_array.append(b_income_after_tax)

#total, bess+solar

for i in years:
	if i == 0: #year 0 has no opex
		t_income_after_tax = np.nan
	else:
		t_income_after_tax = bess_income_after_tax_array[i] + solar_income_after_tax_array[i]
	total_income_after_tax_array.append(t_income_after_tax)

#net cash flow generated the cash flow values for PV BESS and PV+BSS(PCF)

for i in years:
	if i == 0 : #year 0 only takes in initial input values 
		ncf_b = CAPEXbess_array[i]
		ncf_s = CAPEXpv_array[i]
	else: #cash flow for each year after year 0 is the sum of income after tax and CAPEX 
		ncf_b = CAPEXbess_array[i] + bess_income_after_tax_array[i]
		ncf_s = CAPEXpv_array[i] + solar_income_after_tax_array[i]
	pcf = round((ncf_b + ncf_s),2)
	NCF_solar_array.append(ncf_s)
	NCF_bess_array.append(ncf_b)
	PCF_array.append(pcf)

#preparing the arrays to be put into a data frame
CAPEX_list = list(zip(years, CAPEXpv_array, CAPEXbess_array, total_capex, REVpv_array, REVbess_array, totalrev_array, opexpv_array, opexbess_array, totalopex_array, solar_income_array, bess_income_array, solar_tax_total, bess_tax_total, solar_income_after_tax_array, bess_income_after_tax_array, total_income_after_tax_array, NCF_solar_array, NCF_bess_array, PCF_array))
pd.set_option("display.max_rows", None, "display.max_columns", None) #so I can see the entire data frame in sublimes output
df = pd.DataFrame(CAPEX_list, columns = ['year', 'pv capex', 'bess capex', 'total capex', 'PV Revenue', 'bess revenue', 'total revenue','opex PV','opex BESS', 'total OPEX', 'solar income', 'bess income', 'solar tax total', 'bess tax total', 'solar income after tax', 'bess income after tax', 'total income after tax', 'solar ncf', 'bess ncf', 'pcf']) #colunm names for the data frame
print(df)

irr_df = round(np.irr(PCF_array),2) #calculating the IRR  from the projected cash flow colunm in the df
print(irr_df)