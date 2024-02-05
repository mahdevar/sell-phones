from random import sample, shuffle
from itertools import combinations
from time import process_time
from pandas import read_excel, DataFrame, ExcelWriter


PAYS_PER_PURCHASE_PROBABILITY = [100, 20, 5, 2, 1]
SELL_TOLERANCE_LOW = 1.0
SELL_TOLERANCE_HIGH = 1.2
TOLERANCE_INCREASE = 40
#HONOR_PURCHASE_DATES = True
PURCHASE_SPAN = 20
#BASKET_FRACTION = [1, 0.75, 0.5, 0.25]
BASKET_FRACTION = [1, 0.5]

MONTH_START = [0, 31, 62, 93, 124, 155, 187, 217, 247, 277, 307, 337]

# days = lambda date: (date['Year'] - 1400) * 356 + MONTH_START[date['Month'] - 1] + date['Day']
days = lambda date: (int(date[0:4]) - 1402) * 356 + MONTH_START[int(date[5:7]) - 1] + int(date[8:10])

# to_date = lambda date: '%04i/%02i/%02i 12:00:00 ق.ظ' % (date['Year'], date['Month'], date['Day'])
# to_date = lambda date: '%04i/%02i/%02i' % (date['Year'], date['Month'], date['Day'])
# int_column = lambda i: int(i) if i else 0

ALL_CUSTOMERS = DataFrame(read_excel('داده‌ها.xlsx', sheet_name='خریداران', keep_default_na=False, dtype={'کد ملی': str}))
for field in 'سقف', 'خرید', 'شناسه':
	ALL_CUSTOMERS[field].replace('', 0, inplace=True)

OFFICIAL_CUSTOMERS = []
CUSTOMERS = []
for customer in ALL_CUSTOMERS.to_dict('records'):
	if customer['سقف']:
		OFFICIAL_CUSTOMERS += [customer]
	else:
		CUSTOMERS += [customer]


'''
#Remove customers with same name and family
C=[]
UNIQUE_NAMES = set()
for customer in CUSTOMERS:
	if customer['نام'] + customer['نام خانوادگی'] in UNIQUE_NAMES:
		print('[%s]  ----  (%s)' % (customer['نام'], customer['نام خانوادگی']))
	else:
		C += [customer]
		UNIQUE_NAMES.add(customer['نام'] + customer['نام خانوادگی'])

C += OFFICIAL_CUSTOMERS
with ExcelWriter('customers.xlsx') as file:
	RESULT = DataFrame(C)
	for field in 'سقف', 'خرید', 'شناسه':
		RESULT[field].replace(0, '', inplace=True)
	RESULT.to_excel(file, index=False, freeze_panes=(1, 0))

exit()
'''


'''
Remove customers with name equal to family of someone
N = set(customer['نام'] for customer in CUSTOMERS)
F = set(customer['نام خانوادگی'] for customer in CUSTOMERS)

C=[]

for customer in CUSTOMERS:
	if customer['نام خانوادگی'] in N or customer['نام'] in F or '(' in customer['نام'] or '(' in customer['نام خانوادگی'] or len(customer['کد ملی']) != 10:
		pass
		#print('[%s]  ----  (%s)' % (customer['نام'], customer['نام خانوادگی']))
	else:
		C += [customer]

C+=OFFICIAL_CUSTOMERS
with ExcelWriter('customers.xlsx') as file:
	RESULT = DataFrame(C)
	for field in 'سقف', 'خرید', 'شناسه':
		RESULT[field].replace(0, '', inplace=True)
	RESULT.to_excel(file, index=False, freeze_panes=(1, 0))

exit()
'''

PURCHASES = read_excel('داده‌ها.xlsx', sheet_name='گوشی', keep_default_na=False).to_dict('records')
for p in PURCHASES:
	if p['تعداد فروخته شده'] == '-':
		p['تعداد فروخته شده'] = 0
	p['Days'] = days(p['تاريخ'])
PURCHASES = sorted(PURCHASES, key=lambda purchase: purchase['Days'], reverse=True)

#i	0	1	2	3	4	5	6	7	8	9	10
#d	100	100	90	85	20	10	10	10	10	5	1

# p[100]=0
# p[90] = 2
# p[20] = 4

# p[15] = 5

# p[10] = 5


FIRST_PURCHASE_INDEX = [0] * 366
for i, p in enumerate(PURCHASES):
	if not FIRST_PURCHASE_INDEX[p['Days']]:
		FIRST_PURCHASE_INDEX[p['Days']] = i

for i in range(1, 366):
	if not FIRST_PURCHASE_INDEX[i]:
		FIRST_PURCHASE_INDEX[i] = FIRST_PURCHASE_INDEX[i - 1]

DATES = read_excel('داده‌ها.xlsx', sheet_name='واریز', keep_default_na=False).to_dict('records')
DATES = [DATE for DATE in DATES if DATE['بدهکار']]

for DATE in DATES:
	DATE['Days'] = days(DATE['تاریخ'])

# DATES = sample(DATES, k=500)
DATES = sorted(DATES, key=lambda date: date['Days'])


def select_phone(left, date, previous_selection):
	start = FIRST_PURCHASE_INDEX[date['Days']]
	for P in PURCHASES[start:]:
		# TODO: REMOVE THIS UNNECESSARY CONDITION
		if P['Days'] <= date['Days']: # (not HONOR_PURCHASE_DATES) or 
			if P not in previous_selection:
				# for n in 1, 2, 3:
				for n in range(1, min(3, P['مقدار'] - P['تعداد فروخته شده']) + 1):
					for fraction in BASKET_FRACTION:
						percent = left * fraction / (P['مبلغ واحد کالا'] * n)
						if SELL_TOLERANCE_LOW <= percent <= SELL_TOLERANCE_HIGH:
							return {'Phone': P, 'Quantity': n, 'Paid': int(P['مبلغ واحد کالا'] * percent)}
		else:
			print(P['Days'] , date['Days'])
			exit()
			pass  # break
	raise ValueError


ADD = [i for i in range(1, 10)]
for i in range(50):
	ADD += [int(ADD[-9] * 10)]
ADD = sorted(ADD, reverse=True)


def sell(date, money):
	sales = []
	selected_purchases = []
	left = money
	while left != 0:
		try:
			sale = select_phone(left, date, selected_purchases)
		# print(sale[0]['مبلغ واحد کالا'], sale[1] * sale[2])
		except ValueError:
			add_index = 0
			while left > 0 and add_index < len(ADD):
				could_add = False
				for sale in sales:
					if sale['Quantity'] * ADD[add_index] <= left and sale['Paid'] + ADD[add_index] <= sale['Phone']['مبلغ واحد کالا'] * SELL_TOLERANCE_HIGH:
						sale['Paid'] += ADD[add_index]
						left -= sale['Quantity'] * ADD[add_index]
						could_add = True
				if not could_add:
					add_index += 1
			break
		else:
			left -= sale['Quantity'] * sale['Paid']
			sales += [sale]
			selected_purchases += [sale['Phone']]
	return sales, left


def round_sales(sales):
	i = 0
	for S in sales[0:-1]:
		n = S['Paid'] % 1000000
		i += n * S['Quantity']
		S['Paid'] -= n
	sales[-1]['Paid'] += i / sales[-1]['Quantity']
	shuffle(sales)
	return sales


NUMBER_OF_DATES = len(DATES)
LAST_PRINTED_TIME = 0
SALES = []
PAYS = list(range(len(PAYS_PER_PURCHASE_PROBABILITY)))
for tolerance in range(TOLERANCE_INCREASE + 1):
	if TOLERANCE_INCREASE:
		print('TOL:', tolerance)
	s = 0
	while s < len(DATES):
		pays = sample(PAYS, k=1, counts=PAYS_PER_PURCHASE_PROBABILITY)[0]
		if len(DATES) <= s + pays:
			break
		if process_time() > LAST_PRINTED_TIME + 60:
			try:
				print('%im left' % ((len(DATES) - s) * process_time() / (s + (NUMBER_OF_DATES - len(DATES))) / 60))
			except ZeroDivisionError:
				pass
			LAST_PRINTED_TIME = process_time()
		e = s + 1
		while e < len(DATES) and DATES[e]['Days'] < DATES[s]['Days'] + PURCHASE_SPAN:
			e += 1
		for other_dates in combinations(DATES[s + 1: e], pays):
			SELECTED_DATES = (DATES[s],) + other_dates
			paid = sum(date['بدهکار'] for date in SELECTED_DATES)
			SALE, LEFT = sell(SELECTED_DATES[0], paid)
			if LEFT == 0:
				shuffle(OFFICIAL_CUSTOMERS)
				customer = {}
				for customer in OFFICIAL_CUSTOMERS:
					if customer['خرید'] + paid < customer['سقف']:
						customer['خرید'] += paid
						break
				else:
					shuffle(CUSTOMERS)
					for customer in CUSTOMERS:
						if customer['خرید'] == 0 or customer['خرید'] + paid < (customer['سقف'] or 30_000_0000):
							customer['خرید'] += paid
							break
					else:
						print('NO CUSTOMER!!!')
						exit()

				for sale in SALE:
					sale['Phone']['تعداد فروخته شده'] += sale['Quantity']

				SALE = round_sales(SALE)
				SALES += [{'Deals': SALE, 'Dates': SELECTED_DATES, 'Customer': customer}]

				for date in SELECTED_DATES:
					if date in DATES:
						DATES.remove(date)
				print('% .3f' % (sum((sale['Paid'] - sale['Phone']['مبلغ واحد کالا']) * sale['Quantity'] for sale in SALE) / 10000000), len(DATES), ', '.join('%3d: %6.3f' % (date['Days'], date['بدهکار'] / 10000000) for date in SELECTED_DATES), sep='\t')
				break
		else:
			s += 1
	SELL_TOLERANCE_HIGH += 0.03
	SELL_TOLERANCE_LOW -= 0.03


if input('Failed?') == 'y':
	exit()

SALES = sorted(SALES, key=lambda sale: min(date['Days'] for date in sale['Dates']))

CUSTOMERS += OFFICIAL_CUSTOMERS
CUSTOMERS = sorted(CUSTOMERS, key=lambda customer: customer['سقف'], reverse=True)
CUSTOMERS_ID = max([customer['شناسه'] for customer in CUSTOMERS] + [4111111]) + 1

for customer in CUSTOMERS:
	if customer['خرید'] != 0 and customer['شناسه'] == 0:
		customer['شناسه'] = CUSTOMERS_ID
		CUSTOMERS_ID += 1
CUSTOMERS = sorted(CUSTOMERS, key=lambda customer: customer['شناسه'] or 9999999999)


with ExcelWriter('داده‌ها.xlsx', mode='a', if_sheet_exists='replace') as file:
	RESULT = DataFrame(CUSTOMERS)
	for field in 'سقف', 'خرید', 'شناسه':
		RESULT[field].replace(0, '', inplace=True)
	RESULT.to_excel(file, sheet_name='خریداران', index=False, freeze_panes=(1, 0))

	RESULT = DataFrame(DATES)
	if len(RESULT):
		del RESULT['Days']
	RESULT.to_excel(file, sheet_name='واریز', index=False, freeze_panes=(1, 0))

	RESULT = DataFrame(sorted(PURCHASES, key=lambda purchase: purchase['Days']))
	RESULT['تعداد فروخته شده'].replace(0, '-', inplace=True)
	del RESULT['Days']
	RESULT.to_excel(file, sheet_name='گوشی', index=False, freeze_panes=(1, 0))

COLUMNS = DataFrame(read_excel('فاکتور فروش.xlsx')).columns
FACTORS = read_excel('فاکتور فروش.xlsx', keep_default_na=False).to_dict('records')
LAST_FACTOR = FACTORS[-1]['فاكتور شماره'] + 1 if FACTORS else 41113208

for i, sale in enumerate(SALES):
	for p in sale['Deals']:
		RECORD = \
		{
			'فاكتور شماره': i + LAST_FACTOR,
			'فاكتور تاريخ': sale['Dates'][0]['تاریخ'],
			'فاكتور كد مشتري': sale['Customer']['شناسه'],
			'قلم فاكتور كد': p['Phone']['كد كالا'],
			'قلم فاكتور واحد اصلي': p['Quantity'],
			'قلم فاكتور في': p['Paid'] / 1.09,
			'قلم فاكتور كل': p['Quantity'] * (p['Paid'] / 1.09),
			'قلم فاكتور ماليات': p['Quantity'] * (p['Paid'] / 1.09) * 0.09,
			'فاكتور نام مشتري': '%s %s' % (sale['Customer']['نام خانوادگی'], sale['Customer']['نام'])
		}
		FACTORS += [{field: '' for field in COLUMNS} | RECORD]
with ExcelWriter('فاکتور فروش.xlsx', mode='a', if_sheet_exists='overlay') as file:
	DataFrame(FACTORS).to_excel(file, index=False, header=COLUMNS)


COLUMNS = DataFrame(read_excel('رسید دریافت.xlsx')).columns
FACTORS = read_excel('رسید دریافت.xlsx', keep_default_na=False).to_dict('records')
for sale in SALES:
	for date in sale['Dates']:
		RECORD = \
		{
			'رسيد دريافت طرف مقابل': '%s %s' % (sale['Customer']['نام خانوادگی'], sale['Customer']['نام']),
			'رسيد دريافت تاريخ': date['تاریخ'],
			'رسيد دريافت جمع دريافت': date['بدهکار'],
			'حواله تاريخ': date['تاریخ'],
			'حواله مبلغ': date['بدهکار'],
			'نام بانک': date['نام بانک']
			}
		FACTORS += [{field: '' for field in COLUMNS} | RECORD]
with ExcelWriter('رسید دریافت.xlsx', mode='a', if_sheet_exists='overlay') as file:
	DataFrame(FACTORS).to_excel(file, index=False, header=COLUMNS)


COLUMNS = DataFrame(read_excel('طرف حساب.xlsx')).columns
RESULTS = []
for customer in CUSTOMERS:
	if customer['شناسه']:
		RECORD = \
		{
			'طرف حساب نام': customer['نام'],
			'طرف حساب نام خانوادگي': customer['نام خانوادگی'],
			'طرف حساب عنوان': '%s %s' % (customer['نام خانوادگی'], customer['نام']),
			'طرف حساب كدملي/شناسه ملي': customer['کد ملی'],
			'طرف حساب كد': customer['شناسه']
		}
		RESULTS += [{field: '' for field in COLUMNS} | RECORD]
with ExcelWriter('طرف حساب.xlsx') as file:
	DataFrame(RESULTS).to_excel(file, index=False, freeze_panes=(1, 0), header=COLUMNS)
