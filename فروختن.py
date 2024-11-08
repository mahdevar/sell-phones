from random import sample, shuffle
from itertools import combinations
from time import process_time
from pandas import read_excel, DataFrame, ExcelWriter, set_option

LARGE_PAYMENTS_THRESHOLD = 1000_000_000_000_0000
#PAYS_PER_PURCHASE_PROBABILITY = [20, 5, 2, 1]
PAYS_PER_PURCHASE_PROBABILITY = [100, 5, 2, 1]
SELL_TOLERANCE_LOW = 0.99
SELL_TOLERANCE_HIGH = 1.07
TOLERANCE_INCREASE = 100
#HONOR_PURCHASE_DATES = True
PURCHASE_SPAN = 21
MAX_QUANTITY = 3
MULTIPLE_PURCHASE = [100, 10, 5, 1]
DESIRED_FACTOR = 23475 # 0
DESIRED_CUSTOMERS_ID = 4124303


#MULTIPLE_PURCHASE = [100, 2, 1, 0]
#MULTIPLE_PURCHASE = [100, 0, 0, 0]
#CUSTOM_SALES = ['1403/02/23', '1403/02/23', '1403/02/24', '1403/02/27']
CUSTOM_SALES = []
#bank: 106716244000


MONTH_START = [0, 31, 62, 93, 124, 155, 187, 217, 247, 277, 307, 337]

# days = lambda date: (date['Year'] - 1400) * 356 + MONTH_START[date['Month'] - 1] + date['Day']
days = lambda date: (int(date[0:4]) - 1403) * 356 + MONTH_START[int(date[5:7]) - 1] + int(date[8:10])

# to_date = lambda date: '%04i/%02i/%02i 12:00:00 ق.ظ' % (date['Year'], date['Month'], date['Day'])
# to_date = lambda date: '%04i/%02i/%02i' % (date['Year'], date['Month'], date['Day'])
# int_column = lambda i: int(i) if i else 0

#set_option('future.no_silent_downcasting', True)

ALL_CUSTOMERS = DataFrame(read_excel('داده‌ها.xlsx', sheet_name='خریداران', keep_default_na=False, dtype={'کد ملی': str}))
for field in 'سقف', 'خرید', 'شناسه':
	#ALL_CUSTOMERS[field].replace('', 0, inplace=True)
	ALL_CUSTOMERS[field] = ALL_CUSTOMERS[field].replace('', 0)


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
	if p['تعداد فروخته شده'] == '':
		p['تعداد فروخته شده'] = 0
	p['Days'] = days(p['تاريخ'].strip())

PURCHASES = sorted(PURCHASES, key=lambda purchase: purchase['Days'])

#PRIORITY_CODES = (10010012, 10010207)

"""
priority_left = 0
normal_left = 0
priority_price = 0
normal_price = 0

for P in PURCHASES:
	if P['كد كالا'] in PRIORITY_CODES:
		priority_left += P['مقدار']
		priority_price += P['مقدار'] * P['مقدار'] * P['مبلغ واحد کالا']
	else:
		normal_left += P['مقدار']
		normal_price += P['مقدار'] * P['مقدار'] * P['مبلغ واحد کالا']
print(priority_left, normal_left)
print(priority_price, normal_price)

"""

FIRST_PURCHASE_INDEX = [0] * 367
for i, p in enumerate(PURCHASES):
	if not FIRST_PURCHASE_INDEX[p['Days']]:
		FIRST_PURCHASE_INDEX[p['Days']] = i

for i in range(1, 367):
	if not FIRST_PURCHASE_INDEX[i]:
		FIRST_PURCHASE_INDEX[i] = FIRST_PURCHASE_INDEX[i - 1]

DATES = read_excel('داده‌ها.xlsx', sheet_name='واریز', keep_default_na=False).to_dict('records')
DATES = [DATE for DATE in DATES if DATE['بدهکار']]

for DATE in DATES:
	DATE['Days'] = days(DATE['تاریخ'])

#DATES = sample(DATES, k=10)
DATES = sorted(DATES, key=lambda date: date['Days'])


LARGE_PAYMENTS = 0
for DATE in DATES:
	if DATE['بدهکار'] >= LARGE_PAYMENTS_THRESHOLD:
		LARGE_PAYMENTS += DATE['بدهکار']

print('LARGE_PAYMENTS', LARGE_PAYMENTS)


RESERVED_CUSTOMERS = []
while LARGE_PAYMENTS > 0:
	RESERVED_CUSTOMERS += [OFFICIAL_CUSTOMERS[0]]
	LARGE_PAYMENTS -= OFFICIAL_CUSTOMERS[0]['سقف']
	OFFICIAL_CUSTOMERS = OFFICIAL_CUSTOMERS[1:]


'''


SOLD_PREVIOUSLY = read_excel('داده‌ها.xlsx', sheet_name='فروخته شده', keep_default_na=False).to_dict('records')

'''
def select_phone(left, date, previous_selection):
	end = FIRST_PURCHASE_INDEX[date['Days']] + 1
	index = list(range(end))
	shuffle(index)
	for i in index:
		P = PURCHASES[i]
		#if P['كد كالا'] in PRIORITY_CODES and P not in previous_selection:
		if P not in previous_selection:
			# for n in 1, 2, 3:
			for n in range(1, min(MAX_QUANTITY, P['مقدار'] - P['تعداد فروخته شده']) + 1):
				fraction = sample([1, 0.5, 0.33, 0.25], k=1, counts=MULTIPLE_PURCHASE)[0]
				percent = left * fraction / (P['مبلغ واحد کالا'] * n)
				if SELL_TOLERANCE_LOW <= percent <= SELL_TOLERANCE_HIGH:
					return {'Phone': P, 'Quantity': n, 'Paid': int(P['مبلغ واحد کالا'] * percent)}
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
PROFIT = 0
PHONE_PER_PURCHASE = 0


#for p in PURCHASES:
#	print(p)


while len(DATES) and SELL_TOLERANCE_HIGH < 1.75:
	print('[%.2f - %.2f]' % (SELL_TOLERANCE_LOW, SELL_TOLERANCE_HIGH))
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
		if DATES[s]['بدهکار'] >= 80_000_0000:
			pays = 0
		for other_dates in combinations(DATES[s + 1: e], pays):
			SELECTED_DATES = (DATES[s],) + other_dates
			if pays > 0:
				for DATE in SELECTED_DATES:
					if DATE['بدهکار'] >= 80_000_0000:
						break
			paid = sum(date['بدهکار'] for date in SELECTED_DATES)
			SALE, LEFT = sell(SELECTED_DATES[0], paid)
			if LEFT == 0:
				if paid >= LARGE_PAYMENTS_THRESHOLD:
					print('Large:', paid)
					for customer in RESERVED_CUSTOMERS:
						if customer['خرید'] <= customer['سقف']:
							customer['خرید'] += paid
							print('Selling to reserved', pays, customer['نام'])
							break
					else:
						print('aa')
						exit()
				else:
					shuffle(OFFICIAL_CUSTOMERS)
					customer = {}
					for customer in OFFICIAL_CUSTOMERS:
						if customer['خرید'] + paid <= customer['سقف']:
							customer['خرید'] += paid
							break
					else:
						shuffle(CUSTOMERS)
						for customer in CUSTOMERS:
							if customer['خرید'] == 0 or customer['خرید'] + paid <= (customer['سقف'] or 30_000_0000):
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

				priority_left = 0
				normal_left = 0
				PROFIT += sum((sale['Paid'] - sale['Phone']['مبلغ واحد کالا']) * sale['Quantity'] for sale in SALE) / 10000000
				"""
				for P in PURCHASES:
					if P['كد كالا'] in PRIORITY_CODES:
						priority_left += P['مقدار'] - P['تعداد فروخته شده']
					else:
						normal_left += P['مقدار'] - P['تعداد فروخته شده']
				print(priority_left, normal_left, int(PROFIT), end = ' @ ')
				
				if priority_left < 250:
					#tolerance = 100
					#break
					PRIORITY_CODES = set(P['كد كالا'] for P in PURCHASES)
				"""
				print(len(DATES), '[%+4d]' % int(PROFIT), ', '.join('% .2f %s' % (sale['Paid'] / sale['Phone']['مبلغ واحد کالا'] - 1.0, sale['Phone']['عنوان كالا']) for sale in SALE), len(SELECTED_DATES), sep='\t')
				#, ', '.join('%3d: %6.3f' % (date['Days'], date['بدهکار'] / 10000000) for date in SELECTED_DATES), sep='\t')
				break
		else:
			s += 1
	SELL_TOLERANCE_HIGH += 0.01
	SELL_TOLERANCE_LOW -= 0.0075

print('Total profit:', int(PROFIT))


for DATE in CUSTOM_SALES:
	SALES += [{'Deals': [{'Phone': {'عنوان كالا': 'FATEME', 'كد كالا': 'FATEME'}, 'Quantity': 1, 'Paid': 1}], 'Dates': ({'تاریخ': DATE, 'بدهکار': 1, 'نام بانک': 'FATEME', 'Days': days(DATE)},), 'Customer': {'نام': 'FATEME', 'نام خانوادگی': 'FATEME', 'شناسه': -1, }}]
'''
SOLD_PREVIOUSLY = sorted(SOLD_PREVIOUSLY, key=lambda sale: sale['فاكتور شماره'])

for p in SOLD_PREVIOUSLY:
	print(p)

for sale in SOLD_PREVIOUSLY:
	record = \
	{
		'فاكتور تاريخ': sale['Dates'][0]['تاریخ'],
		'فاكتور كد مشتري': sale['Customer']['شناسه'],
		'قلم فاكتور كد': p['Phone']['كد كالا'],
		'قلم فاكتور واحد اصلي': p['Quantity'],
		'قلم فاكتور في': p['Paid'] / 1.10,
		'قلم فاكتور كل': p['Quantity'] * (p['Paid'] / 1.10),
		'قلم فاكتور ماليات': p['Quantity'] * (p['Paid'] / 1.10) * 0.10,
		'فاكتور نام مشتري': '%s %s' % (sale['Customer']['نام خانوادگی'], sale['Customer']['نام'])
	}

'''



SALES = sorted(SALES, key=lambda sale: min(date['Days'] for date in sale['Dates']))





if input('Failed?') == 'y':
	exit()


CUSTOMERS += OFFICIAL_CUSTOMERS + RESERVED_CUSTOMERS
CUSTOMERS = sorted(CUSTOMERS, key=lambda customer: customer['سقف'], reverse=True)
CUSTOMERS_ID = max([customer['شناسه'] for customer in CUSTOMERS] + [4111111]) + 1

if DESIRED_CUSTOMERS_ID:
	CUSTOMERS_ID = DESIRED_CUSTOMERS_ID

for customer in CUSTOMERS:
	if customer['خرید'] != 0 and customer['شناسه'] == 0:
		customer['شناسه'] = CUSTOMERS_ID
		CUSTOMERS_ID += 1
CUSTOMERS = sorted(CUSTOMERS, key=lambda customer: customer['شناسه'] or 9999999999)


with ExcelWriter('داده‌ها.xlsx', mode='a', if_sheet_exists='replace') as file:
	RESULT = DataFrame(CUSTOMERS)
	for field in 'سقف', 'خرید', 'شناسه':
		#RESULT[field].replace(0, '', inplace=True)
		RESULT[field] = RESULT[field].replace(0, '')
	RESULT.to_excel(file, sheet_name='خریداران', index=False, freeze_panes=(1, 0))

	RESULT = DataFrame(DATES)
	if len(RESULT):
		del RESULT['Days']
	RESULT.to_excel(file, sheet_name='واریز', index=False, freeze_panes=(1, 0))

	RESULT = DataFrame(sorted(PURCHASES, key=lambda purchase: purchase['Days']))
	#RESULT['تعداد فروخته شده'].replace(0, '-', inplace=True)
	#RESULT['تعداد فروخته شده'] = RESULT['تعداد فروخته شده'].replace(0, '-')

	del RESULT['Days']
	RESULT.to_excel(file, sheet_name='گوشی', index=False, freeze_panes=(1, 0))

COLUMNS = DataFrame(read_excel('فاکتور فروش.xlsx')).columns
FACTORS = read_excel('فاکتور فروش.xlsx', keep_default_na=False).to_dict('records')
if DESIRED_FACTOR:
	LAST_FACTOR = DESIRED_FACTOR
else:
	LAST_FACTOR = FACTORS[-1]['فاكتور شماره'] + 1 if FACTORS else 21266

for i, sale in enumerate(SALES):
	for p in sale['Deals']:
		RECORD = \
		{
			'فاكتور شماره': i + LAST_FACTOR,
			'فاكتور تاريخ': sale['Dates'][0]['تاریخ'],
			'فاكتور كد مشتري': sale['Customer']['شناسه'],
			'قلم فاكتور كد': p['Phone']['كد كالا'],
			'قلم فاكتور واحد اصلي': p['Quantity'],
			'قلم فاكتور في': p['Paid'] / 1.10,
			'قلم فاكتور كل': p['Quantity'] * (p['Paid'] / 1.10),
			'قلم فاكتور ماليات': p['Quantity'] * (p['Paid'] / 1.10) * 0.10,
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
