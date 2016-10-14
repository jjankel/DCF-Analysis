from numpy import pv,fv,pmt
from openpyxl import Workbook
import xlsxwriter
import sendtext


"GENERATING YEARS IN PROJECTION "
# pastnum = past periods
# futurenum = Future number of periods
#presper = the present year
compname = input('Please input the name of the company: ')
pastnum = int(input('input the number of past periods: '))
futurenum = int(input('input the number of future periods: '))
presper = int(input('input the present year: '))
interest = float(input('input the current interest rate: '))
print()

"BUILDING THE YEARS LIST FOR THE TABLE"
# years = list of uears in the projection
years = []
for year in range (presper - pastnum, presper + futurenum + 1, 1):
    years.append(year)

"USE N TO GENERATE NUMBER OF CASHFLOW ENTRIES"

cnt = pastnum
Revenues = []
OrigRevenues = []
Opcosts = []
Taxes = []
Netinvests = []
Workcaps = []

print ('Paramaters for Calculating Cashflows are:')
print('Revenues, Operating Costs, Net Investments, Change in Working Capitol')
print()

for past in range(0,cnt + 1,1):
    Revenues.append(float(input('input Revenue from year ' + str(years[past]) + ': ')))
    Opcosts.append(float(input('input Operating Costs from year ' + str(years[past]) + ': ')))
    Taxes.append(float(input('input Taxes from year ' + str(years[past]) + ': ')))
    Netinvests.append(float(input('input Net Investment from year ' + str(years[past]) + ': ')))
    Workcaps.append(float(input('input Change in Working cap from year ' + str(years[past]) + ': ')))
    print()

for i in Revenues:
    OrigRevenues.append(i)

"GENERATE THE REST OF THE VARIABLES"

print("""Would you like to use multiplier method or WACC method for calculating the fair value of equity?""")
print('Specify with: \'multi\' or \'wacc\'')
evaltype = str(input())

"EXCEL"
ALPHA = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
count = 0
spacing = futurenum + pastnum
wb = xlsxwriter.Workbook(compname + 'DCF' + '.xlsx')
ws = wb.add_worksheet('Realistic')
ws2 = wb.add_worksheet('Optomistic')  # We can copy code over when done
ws3 = wb.add_worksheet('Pessimistic')  # We can copy code over when done
"FORMATTING OPTIONS"
chart = wb.add_chart({'type': 'column'})
chart2 = wb.add_chart({'type': 'column'})
chart3 = wb.add_chart({'type': 'column'})
bold = wb.add_format({'bold': True})
grey = wb.add_format()
grey.set_bg_color('gray')
yellow = wb.add_format()
yellow.set_bg_color('yellow')
blue = wb.add_format()
blue.set_bg_color('blue')
money = wb.add_format()
money.set_num_format('$#,##0.00;[Red]($#,##0.00)')
yelmon = wb.add_format()
yelmon.set_bg_color('yellow')
yelmon.set_num_format('$#,##0.00;[Red]($#,##0.00)')

if evaltype.lower() == 'multi':
    Multiplier = int(input('Please Input your Multiplier: '))
    print()
    Debt = float(input('Please input your outstanding debt in the year ' + str(presper) + ': '))
    "GENERATE ALL REVENUE PROJ"
    optomistic = []
    realistic = []
    pessimistic = []
    revgrowth = float(input('input revenue growth (decimal): '))
    opgrowth = revgrowth
    for co in range(0, futurenum, 1):
        optomistic.append(opgrowth)
        opgrowth = opgrowth * (.99)
    regrowth = revgrowth
    for co in range(0, futurenum, 1):
        realistic.append(regrowth)
        regrowth = regrowth * (.97)
    pesgrowth = revgrowth
    for co in range(0, futurenum, 1):
        pessimistic.append(pesgrowth)
        pesgrowth = pesgrowth * (.94)
    "GROWTH OF EXPENSES"
    opcostrt = float(input('input growth rate of operating costs: '))
    taxrt = float(input('input growth rate of taxes: '))
    netinvrt = float(input('input growth rate of net investments: '))
    chngcaprt = float(input('input growth rate of change in working capitol: '))
    print()
    "_______________WS1________________"
    "SHEET SETUP"
    ws.write('A3', 'Revenues', bold)
    ws.write('A5', 'Operating Costs', bold)
    ws.write('A6', 'Taxes', bold)
    ws.write('A7', 'Net Investments', bold)
    ws.write('A8', 'Change in Working Capitol', bold)
    ws.write('A10', 'Cashflow', bold)
    ws.write('A11', '(Debt)', bold)
    ws.write('A13', 'Multiplier', bold)
    ws.write('A14', 'Terminal Value', bold)
    ws.write('A15', 'Enterprise Value', bold)
    ws.write('A16', 'Fair Value of Equity', bold)
    ws.write('A18', 'Revenues', bold)
    ws.write('A19', 'Optomistic', bold)
    ws.write('A20', 'Realistic', bold)
    ws.write('A21', 'Pessimistic', bold)
    for i in range(0, pastnum + 1, 1):
        count = count + 1
        ws.write(ALPHA[count] + '3', Revenues[count - 1], money)
        ws.write(ALPHA[count] + '5', Opcosts[count - 1], money)
        ws.write(ALPHA[count] + '6', Taxes[count - 1], money)
        ws.write(ALPHA[count] + '7', Netinvests[count - 1], money)
        ws.write(ALPHA[count] + '8', Workcaps[count - 1], money)
    for i in range(0, pastnum + futurenum + 1, 1):
        ws.write(ALPHA[i + 1] + '2', years[i], bold)
    ws.write(ALPHA[spacing + 3] + '2', 'Cost Of Equity', bold)
    ws.write(ALPHA[spacing + 3] + '3', 'Cost of Debt', bold)
    ws.write(ALPHA[spacing + 3] + '4', 'beta', bold)
    ws.write(ALPHA[spacing + 3] + '5', 'Rf', bold)
    ws.write(ALPHA[spacing + 3] + '6', 'Rm', bold)
    ws.write(ALPHA[spacing + 3] + '7', 'Re', bold)
    ws.write(ALPHA[spacing + 3] + '8', 'E/V', bold)
    ws.write(ALPHA[spacing + 3] + '9', 'D/V', bold)
    ws.write(ALPHA[spacing + 3] + '10', 'Corp Tax Rate', bold)
    ws.write(ALPHA[spacing + 3] + '11', 'WACC', bold)
    ws.write(ALPHA[spacing + 4] + '15', compname, bold)
    "EQUITY VALUES"
    ws.write(ALPHA[spacing + 4] + '2', '-')
    ws.write(ALPHA[spacing + 4] + '3', '-')
    ws.write(ALPHA[spacing + 4] + '4', '-')
    ws.write(ALPHA[spacing + 4] + '5', '-')
    ws.write(ALPHA[spacing + 4] + '6', '-')
    ws.write(ALPHA[spacing + 4] + '7', '-')
    ws.write(ALPHA[spacing + 4] + '8', '-')
    ws.write(ALPHA[spacing + 4] + '9', '-')
    ws.write(ALPHA[spacing + 4] + '10', '-')
    ws.write(ALPHA[spacing + 4] + '11', '-')
    "REVENUES"
    for i in range(0, futurenum, 1):
        ws.write(ALPHA[i + 2 + pastnum] + '3', (round((1 + realistic[i]) * Revenues[pastnum + i], 2)), money)
        Revenues.append((1 + realistic[i]) * Revenues[pastnum + i])
        ws.write(ALPHA[i + 1] + '20', round((1 + realistic[i]) * Revenues[pastnum + i], 2), money)
    "REVENUE PROJECTION NUMBERS"
    for k in range(0, futurenum, 1):
        ws.write(ALPHA[k + 1] + '19', round((1 + optomistic[k]) * Revenues[pastnum + k], 2), money)
        ws.write(ALPHA[k + 1] + '21', round((1 + pessimistic[k]) * Revenues[pastnum + k], 2), money)
    "THE REST"
    for i in range(0, futurenum, 1):
        ws.write(ALPHA[i + 2 + pastnum] + '5', round((1 + opcostrt) * Opcosts[pastnum + i], 2), money)
        Opcosts.append((1 + opcostrt) * Opcosts[pastnum + i])
        ws.write(ALPHA[i + 2 + pastnum] + '6', round((1 + taxrt) * Taxes[pastnum + i], 2), money)
        Taxes.append((1 + taxrt) * Taxes[pastnum + i])
        ws.write(ALPHA[i + 2 + pastnum] + '7', round((1 + netinvrt) * Netinvests[pastnum + i], 2), money)
        Netinvests.append((1 + netinvrt) * Netinvests[pastnum + i])
        ws.write(ALPHA[i + 2 + pastnum] + '8', round((1 + chngcaprt) * Workcaps[pastnum + i], 2), money)
        Workcaps.append((1 + chngcaprt) * Workcaps[pastnum + i])
    "CASHFLOWS"
    Cashflow = []
    for y in range(0, spacing + 1, 1):
        ws.write(ALPHA[y + 1] + '10', round(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]), 2),
                 money)
        Cashflow.append(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]))
    "DEBT & MULTIPLIER"
    ws.write(ALPHA[2 + pastnum] + '11', Debt, money)
    ws.write('B13', '(' + str(Multiplier) + 'X' + ')')
    "TERMINAL, ENTERPRISE, FAIR VAL OF EQUITY"
    Fincf = Cashflow[spacing]
    Termval = Fincf * Multiplier
    Entvalue = -(pv(interest, futurenum, 0, Termval))
    Fairval = Entvalue - Debt
    ws.write('B14', Termval, money)
    ws.write('B15', Entvalue, money)
    ws.write('B16', Fairval, yelmon)
    "_______________WS2________________"
    "SHEET SETUP"
    ws2.write('A3', 'Revenues', bold)
    ws2.write('A5', 'Operating Costs', bold)
    ws2.write('A6', 'Taxes', bold)
    ws2.write('A7', 'Net Investments', bold)
    ws2.write('A8', 'Change in Working Capitol', bold)
    ws2.write('A10', 'Cashflow', bold)
    ws2.write('A11', '(Debt)', bold)
    ws2.write('A13', 'Multiplier', bold)
    ws2.write('A14', 'Terminal Value', bold)
    ws2.write('A15', 'Enterprise Value', bold)
    ws2.write('A16', 'Fair Value of Equity', bold)
    ws2.write('A18', 'Revenues', bold)
    ws2.write('A19', 'Optomistic', bold)
    ws2.write('A20', 'Realistic', bold)
    ws2.write('A21', 'Pessimistic', bold)
    count = 0
    for i in range(0, pastnum + 1, 1):
        count = count + 1
        ws2.write(ALPHA[count] + '3', Revenues[count - 1], money)
        ws2.write(ALPHA[count] + '5', Opcosts[count - 1], money)
        ws2.write(ALPHA[count] + '6', Taxes[count - 1], money)
        ws2.write(ALPHA[count] + '7', Netinvests[count - 1], money)
        ws2.write(ALPHA[count] + '8', Workcaps[count - 1], money)
    for i in range(0, pastnum + futurenum + 1, 1):
        ws2.write(ALPHA[i + 1] + '2', years[i], bold)
    ws2.write(ALPHA[spacing + 3] + '2', 'Cost Of Equity', bold)
    ws2.write(ALPHA[spacing + 3] + '3', 'Cost of Debt', bold)
    ws2.write(ALPHA[spacing + 3] + '4', 'beta', bold)
    ws2.write(ALPHA[spacing + 3] + '5', 'Rf', bold)
    ws2.write(ALPHA[spacing + 3] + '6', 'Rm', bold)
    ws2.write(ALPHA[spacing + 3] + '7', 'Re', bold)
    ws2.write(ALPHA[spacing + 3] + '8', 'E/V', bold)
    ws2.write(ALPHA[spacing + 3] + '9', 'D/V', bold)
    ws2.write(ALPHA[spacing + 3] + '10', 'Corp Tax Rate', bold)
    ws2.write(ALPHA[spacing + 3] + '11', 'WACC', bold)
    ws2.write(ALPHA[spacing + 4] + '15', compname, bold)
    "EQUITY VALUES"
    ws2.write(ALPHA[spacing + 4] + '2', '-')
    ws2.write(ALPHA[spacing + 4] + '3', '-')
    ws2.write(ALPHA[spacing + 4] + '4', '-')
    ws2.write(ALPHA[spacing + 4] + '5', '-')
    ws2.write(ALPHA[spacing + 4] + '6', '-')
    ws2.write(ALPHA[spacing + 4] + '7', '-')
    ws2.write(ALPHA[spacing + 4] + '8', '-')
    ws2.write(ALPHA[spacing + 4] + '9', '-')
    ws2.write(ALPHA[spacing + 4] + '10', '-')
    ws2.write(ALPHA[spacing + 4] + '11', '-')
    "REVENUES"
    Revenues[:] = []
    print(OrigRevenues)
    for i in OrigRevenues:
        Revenues.append(i)
    for i in range(0, futurenum, 1):
        ws2.write(ALPHA[i + 2 + pastnum] + '3', (round((1 + optomistic[i]) * Revenues[pastnum + i], 2)), money)
        Revenues.append((1 + optomistic[i]) * Revenues[pastnum + i])
        ws2.write(ALPHA[i + 1] + '19', round((1 + optomistic[i]) * Revenues[pastnum + i], 2), money)
    "REVENUE PROJECTION NUMBERS"
    for k in range(0, futurenum, 1):
        ws2.write(ALPHA[k + 1] + '20', round((1 + realistic[k]) * Revenues[pastnum + k], 2), money)
        ws2.write(ALPHA[k + 1] + '21', round((1 + pessimistic[k]) * Revenues[pastnum + k], 2), money)
    "THE REST"
    for i in range(0, futurenum, 1):
        ws2.write(ALPHA[i + 2 + pastnum] + '5', round((1 + opcostrt) * Opcosts[pastnum + i], 2), money)
        Opcosts.append((1 + opcostrt) * Opcosts[pastnum + i])
        ws2.write(ALPHA[i + 2 + pastnum] + '6', round((1 + taxrt) * Taxes[pastnum + i], 2), money)
        Taxes.append((1 + taxrt) * Taxes[pastnum + i])
        ws2.write(ALPHA[i + 2 + pastnum] + '7', round((1 + netinvrt) * Netinvests[pastnum + i], 2), money)
        Netinvests.append((1 + netinvrt) * Netinvests[pastnum + i])
        ws2.write(ALPHA[i + 2 + pastnum] + '8', round((1 + chngcaprt) * Workcaps[pastnum + i], 2), money)
        Workcaps.append((1 + chngcaprt) * Workcaps[pastnum + i])
    "CASHFLOWS"
    Cashflow[:] = []
    for y in range(0, spacing + 1, 1):
        ws2.write(ALPHA[y + 1] + '10', round(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]), 2),
                  money)
        Cashflow.append(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]))
    "DEBT & MULTIPLIER"
    ws2.write(ALPHA[2 + pastnum] + '11', Debt, money)
    ws2.write('B13', '(' + str(Multiplier) + 'X' + ')')
    "TERMINAL, ENTERPRISE, FAIR VAL OF EQUITY"
    Fincf = Cashflow[spacing]
    Termval = Fincf * Multiplier
    Entvalue = -(pv(interest, futurenum, 0, Termval))
    Fairval = Entvalue - Debt
    ws2.write('B14', Termval, money)
    ws2.write('B15', Entvalue, money)
    ws2.write('B16', Fairval, yelmon)
    "_______________WS3________________"
    "SHEET SETUP"
    ws3.write('A3', 'Revenues', bold)
    ws3.write('A5', 'Operating Costs', bold)
    ws3.write('A6', 'Taxes', bold)
    ws3.write('A7', 'Net Investments', bold)
    ws3.write('A8', 'Change in Working Capitol', bold)
    ws3.write('A10', 'Cashflow', bold)
    ws3.write('A11', '(Debt)', bold)
    ws3.write('A13', 'Multiplier', bold)
    ws3.write('A14', 'Terminal Value', bold)
    ws3.write('A15', 'Enterprise Value', bold)
    ws3.write('A16', 'Fair Value of Equity', bold)
    ws3.write('A18', 'Revenues', bold)
    ws3.write('A19', 'Optomistic', bold)
    ws3.write('A20', 'Realistic', bold)
    ws3.write('A21', 'Pessimistic', bold)
    count = 0
    for i in range(0, pastnum + 1, 1):
        count = count + 1
        ws3.write(ALPHA[count] + '3', Revenues[count - 1], money)
        ws3.write(ALPHA[count] + '5', Opcosts[count - 1], money)
        ws3.write(ALPHA[count] + '6', Taxes[count - 1], money)
        ws3.write(ALPHA[count] + '7', Netinvests[count - 1], money)
        ws3.write(ALPHA[count] + '8', Workcaps[count - 1], money)
    for i in range(0, pastnum + futurenum + 1, 1):
        ws3.write(ALPHA[i + 1] + '2', years[i], bold)
    ws3.write(ALPHA[spacing + 3] + '2', 'Cost Of Equity', bold)
    ws3.write(ALPHA[spacing + 3] + '3', 'Cost of Debt', bold)
    ws3.write(ALPHA[spacing + 3] + '4', 'beta', bold)
    ws3.write(ALPHA[spacing + 3] + '5', 'Rf', bold)
    ws3.write(ALPHA[spacing + 3] + '6', 'Rm', bold)
    ws3.write(ALPHA[spacing + 3] + '7', 'Re', bold)
    ws3.write(ALPHA[spacing + 3] + '8', 'E/V', bold)
    ws3.write(ALPHA[spacing + 3] + '9', 'D/V', bold)
    ws3.write(ALPHA[spacing + 3] + '10', 'Corp Tax Rate', bold)
    ws3.write(ALPHA[spacing + 3] + '11', 'WACC', bold)
    ws3.write(ALPHA[spacing + 4] + '15', compname, bold)
    "EQUITY VALUES"
    ws3.write(ALPHA[spacing + 4] + '2', '-')
    ws3.write(ALPHA[spacing + 4] + '3', '-')
    ws3.write(ALPHA[spacing + 4] + '4', '-')
    ws3.write(ALPHA[spacing + 4] + '5', '-')
    ws3.write(ALPHA[spacing + 4] + '6', '-')
    ws3.write(ALPHA[spacing + 4] + '7', '-')
    ws3.write(ALPHA[spacing + 4] + '8', '-')
    ws3.write(ALPHA[spacing + 4] + '9', '-')
    ws3.write(ALPHA[spacing + 4] + '10', '-')
    ws3.write(ALPHA[spacing + 4] + '11', '-')
    "REVENUES"
    Revenues[:] = []
    print(OrigRevenues)
    for i in OrigRevenues:
        Revenues.append(i)
    for i in range(0, futurenum, 1):
        ws3.write(ALPHA[i + 2 + pastnum] + '3', (round((1 + pessimistic[i]) * Revenues[pastnum + i], 2)), money)
        Revenues.append((1 + pessimistic[i]) * Revenues[pastnum + i])
        ws3.write(ALPHA[i + 1] + '21', round((1 + pessimistic[i]) * Revenues[pastnum + i], 2), money)
    "REVENUE PROJECTION NUMBERS"
    for k in range(0, futurenum, 1):
        ws3.write(ALPHA[k + 1] + '20', round((1 + realistic[k]) * Revenues[pastnum + k], 2), money)
        ws3.write(ALPHA[k + 1] + '19', round((1 + optomistic[k]) * Revenues[pastnum + k], 2), money)
    "THE REST"
    for i in range(0, futurenum, 1):
        ws3.write(ALPHA[i + 2 + pastnum] + '5', round((1 + opcostrt) * Opcosts[pastnum + i], 2), money)
        Opcosts.append((1 + opcostrt) * Opcosts[pastnum + i])
        ws3.write(ALPHA[i + 2 + pastnum] + '6', round((1 + taxrt) * Taxes[pastnum + i], 2), money)
        Taxes.append((1 + taxrt) * Taxes[pastnum + i])
        ws3.write(ALPHA[i + 2 + pastnum] + '7', round((1 + netinvrt) * Netinvests[pastnum + i], 2), money)
        Netinvests.append((1 + netinvrt) * Netinvests[pastnum + i])
        ws3.write(ALPHA[i + 2 + pastnum] + '8', round((1 + chngcaprt) * Workcaps[pastnum + i], 2), money)
        Workcaps.append((1 + chngcaprt) * Workcaps[pastnum + i])
    "CASHFLOWS"
    Cashflow[:] = []
    for y in range(0, spacing + 1, 1):
        ws3.write(ALPHA[y + 1] + '10', round(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]), 2),
                  money)
        Cashflow.append(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]))
    "DEBT & MULTIPLIER"
    ws3.write(ALPHA[2 + pastnum] + '11', Debt, money)
    ws3.write('B13', '(' + str(Multiplier) + 'X' + ')')
    "TERMINAL, ENTERPRISE, FAIR VAL OF EQUITY"
    Fincf = Cashflow[spacing]
    Termval = Fincf * Multiplier
    Entvalue = -(pv(interest, futurenum, 0, Termval))
    Fairval = Entvalue - Debt
    ws3.write('B14', Termval, money)
    ws3.write('B15', Entvalue, money)
    ws3.write('B16', Fairval, yelmon)
elif evaltype.lower() == 'wacc':
    beta = float(input('Please input beta: '))
    Rf = float(input('Please input Rf: '))
    Rm = float(input('Please input Rm: '))
    Re = float(input('Please input Re: '))
    EV = float(input('Please input E/V: '))
    DV = float(input('Please input D/V: '))
    Corptaxrate = float(input('Please input Corporate Tax Rate: '))
    Costofeq = Rf + beta * (Rm - Rf)
    Costofde = float(input('Please input the cost of debt: '))
    WACC = (Re * EV) + (Costofde * (1 - Corptaxrate) * DV)
    print()
    Debt = float(input('Please input your outstanding debt in the year ' + str(presper) + ': '))
    "GENERATE ALL REVENUE PROJ"
    optomistic = []
    realistic = []
    pessimistic = []
    revgrowth = float(input('input revenue growth (decimal): '))
    opgrowth = revgrowth
    for co in range(0, futurenum, 1):
        optomistic.append(opgrowth)
        opgrowth = opgrowth * (.99)
    regrowth = revgrowth
    for co in range(0, futurenum, 1):
        realistic.append(regrowth)
        regrowth = regrowth * (.97)
    pesgrowth = revgrowth
    for co in range(0, futurenum, 1):
        pessimistic.append(pesgrowth)
        pesgrowth = pesgrowth * (.94)
    "GROWTH OF EXPENSES"
    opcostrt = float(input('input growth rate of operating costs: '))
    taxrt = float(input('input growth rate of taxes: '))
    netinvrt = float(input('input growth rate of net investments: '))
    chngcaprt = float(input('input growth rate of change in working capitol: '))
    print()
    "_______________WS1________________"
    "SHEET SETUP"
    ws.write('A3', 'Revenues', bold)
    ws.write('A5', 'Operating Costs', bold)
    ws.write('A6', 'Taxes', bold)
    ws.write('A7', 'Net Investments', bold)
    ws.write('A8', 'Change in Working Capitol', bold)
    ws.write('A10', 'Cashflow', bold)
    ws.write('A11', '(Debt)', bold)
    ws.write('A13', 'Multiplier', bold)
    ws.write('A14', 'Terminal Value', bold)
    ws.write('A15', 'Enterprise Value', bold)
    ws.write('A16', 'Fair Value of Equity', bold)
    ws.write('A18', 'Revenues', bold)
    ws.write('A19', 'Optomistic', bold)
    ws.write('A20', 'Realistic', bold)
    ws.write('A21', 'Pessimistic', bold)
    for i in range(0, pastnum + 1, 1):
        count = count + 1
        ws.write(ALPHA[count] + '3', Revenues[count - 1], money)
        ws.write(ALPHA[count] + '5', Opcosts[count - 1], money)
        ws.write(ALPHA[count] + '6', Taxes[count - 1], money)
        ws.write(ALPHA[count] + '7', Netinvests[count - 1], money)
        ws.write(ALPHA[count] + '8', Workcaps[count - 1], money)
    for i in range(0, pastnum + futurenum + 1, 1):
        ws.write(ALPHA[i + 1] + '2', years[i], bold)
    ws.write(ALPHA[spacing + 3] + '2', 'Cost Of Equity', bold)
    ws.write(ALPHA[spacing + 3] + '3', 'Cost of Debt', bold)
    ws.write(ALPHA[spacing + 3] + '4', 'beta', bold)
    ws.write(ALPHA[spacing + 3] + '5', 'Rf', bold)
    ws.write(ALPHA[spacing + 3] + '6', 'Rm', bold)
    ws.write(ALPHA[spacing + 3] + '7', 'Re', bold)
    ws.write(ALPHA[spacing + 3] + '8', 'E/V', bold)
    ws.write(ALPHA[spacing + 3] + '9', 'D/V', bold)
    ws.write(ALPHA[spacing + 3] + '10', 'Corp Tax Rate', bold)
    ws.write(ALPHA[spacing + 3] + '11', 'WACC', bold)
    ws.write(ALPHA[spacing + 4] + '15', compname, bold)
    "EQUITY VALUES"
    ws.write(ALPHA[spacing + 4] + '2', Costofeq)
    ws.write(ALPHA[spacing + 4] + '3', Costofde)
    ws.write(ALPHA[spacing + 4] + '4', beta)
    ws.write(ALPHA[spacing + 4] + '5', Rf)
    ws.write(ALPHA[spacing + 4] + '6', Rm)
    ws.write(ALPHA[spacing + 4] + '7', Re)
    ws.write(ALPHA[spacing + 4] + '8', EV)
    ws.write(ALPHA[spacing + 4] + '9', DV)
    ws.write(ALPHA[spacing + 4] + '10', Corptaxrate)
    ws.write(ALPHA[spacing + 4] + '11', WACC)
    "REVENUES"
    for i in range(0, futurenum, 1):
        ws.write(ALPHA[i + 2 + pastnum] + '3', (round((1 + realistic[i]) * Revenues[pastnum + i], 2)), money)
        Revenues.append((1 + realistic[i]) * Revenues[pastnum + i])
        ws.write(ALPHA[i + 1] + '20', round((1 + realistic[i]) * Revenues[pastnum + i], 2), money)
    "REVENUE PROJECTION NUMBERS"
    for k in range(0, futurenum, 1):
        ws.write(ALPHA[k + 1] + '19', round((1 + optomistic[k]) * Revenues[pastnum + k], 2), money)
        ws.write(ALPHA[k + 1] + '21', round((1 + pessimistic[k]) * Revenues[pastnum + k], 2), money)
    "THE REST"
    for i in range(0, futurenum, 1):
        ws.write(ALPHA[i + 2 + pastnum] + '5', round((1 + opcostrt) * Opcosts[pastnum + i], 2), money)
        Opcosts.append((1 + opcostrt) * Opcosts[pastnum + i])
        ws.write(ALPHA[i + 2 + pastnum] + '6', round((1 + taxrt) * Taxes[pastnum + i], 2), money)
        Taxes.append((1 + taxrt) * Taxes[pastnum + i])
        ws.write(ALPHA[i + 2 + pastnum] + '7', round((1 + netinvrt) * Netinvests[pastnum + i], 2), money)
        Netinvests.append((1 + netinvrt) * Netinvests[pastnum + i])
        ws.write(ALPHA[i + 2 + pastnum] + '8', round((1 + chngcaprt) * Workcaps[pastnum + i], 2), money)
        Workcaps.append((1 + chngcaprt) * Workcaps[pastnum + i])
    "CASHFLOWS"
    Cashflow = []
    for y in range(0, spacing + 1, 1):
        ws.write(ALPHA[y + 1] + '10', round(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]), 2),
                 money)
        Cashflow.append(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]))
    "DEBT & MULTIPLIER"
    ws.write(ALPHA[2 + pastnum] + '11', Debt, money)
    ws.write('B13', '-')
    "TERMINAL, ENTERPRISE, FAIR VAL OF EQUITY"
    Fincf = Cashflow[spacing]
    Termval = Fincf * ((1 + revgrowth) / (WACC - revgrowth))
    Entvalue = -(pv(interest, futurenum, 0, Termval))
    Fairval = Entvalue - Debt
    ws.write('B14', Termval, money)
    ws.write('B15', Entvalue, money)
    ws.write('B16', Fairval, yelmon)
    "_______________WS2________________"
    "SHEET SETUP"
    ws2.write('A3', 'Revenues', bold)
    ws2.write('A5', 'Operating Costs', bold)
    ws2.write('A6', 'Taxes', bold)
    ws2.write('A7', 'Net Investments', bold)
    ws2.write('A8', 'Change in Working Capitol', bold)
    ws2.write('A10', 'Cashflow', bold)
    ws2.write('A11', '(Debt)', bold)
    ws2.write('A13', 'Multiplier', bold)
    ws2.write('A14', 'Terminal Value', bold)
    ws2.write('A15', 'Enterprise Value', bold)
    ws2.write('A16', 'Fair Value of Equity', bold)
    ws2.write('A18', 'Revenues', bold)
    ws2.write('A19', 'Optomistic', bold)
    ws2.write('A20', 'Realistic', bold)
    ws2.write('A21', 'Pessimistic', bold)
    count = 0
    for i in range(0, pastnum + 1, 1):
        count = count + 1
        ws2.write(ALPHA[count] + '3', Revenues[count - 1], money)
        ws2.write(ALPHA[count] + '5', Opcosts[count - 1], money)
        ws2.write(ALPHA[count] + '6', Taxes[count - 1], money)
        ws2.write(ALPHA[count] + '7', Netinvests[count - 1], money)
        ws2.write(ALPHA[count] + '8', Workcaps[count - 1], money)
    for i in range(0, pastnum + futurenum + 1, 1):
        ws2.write(ALPHA[i + 1] + '2', years[i], bold)
    ws2.write(ALPHA[spacing + 3] + '2', 'Cost Of Equity', bold)
    ws2.write(ALPHA[spacing + 3] + '3', 'Cost of Debt', bold)
    ws2.write(ALPHA[spacing + 3] + '4', 'beta', bold)
    ws2.write(ALPHA[spacing + 3] + '5', 'Rf', bold)
    ws2.write(ALPHA[spacing + 3] + '6', 'Rm', bold)
    ws2.write(ALPHA[spacing + 3] + '7', 'Re', bold)
    ws2.write(ALPHA[spacing + 3] + '8', 'E/V', bold)
    ws2.write(ALPHA[spacing + 3] + '9', 'D/V', bold)
    ws2.write(ALPHA[spacing + 3] + '10', 'Corp Tax Rate', bold)
    ws2.write(ALPHA[spacing + 3] + '11', 'WACC', bold)
    ws2.write(ALPHA[spacing + 4] + '15', compname, bold)
    "EQUITY VALUES"
    ws2.write(ALPHA[spacing + 4] + '2', Costofeq)
    ws2.write(ALPHA[spacing + 4] + '3', Costofde)
    ws2.write(ALPHA[spacing + 4] + '4', beta)
    ws2.write(ALPHA[spacing + 4] + '5', Rf)
    ws2.write(ALPHA[spacing + 4] + '6', Rm)
    ws2.write(ALPHA[spacing + 4] + '7', Re)
    ws2.write(ALPHA[spacing + 4] + '8', EV)
    ws2.write(ALPHA[spacing + 4] + '9', DV)
    ws2.write(ALPHA[spacing + 4] + '10', Corptaxrate)
    ws2.write(ALPHA[spacing + 4] + '11', WACC)
    "REVENUES"
    Revenues[:] = []
    for i in OrigRevenues:
        Revenues.append(i)
    for i in range(0, futurenum, 1):
        ws2.write(ALPHA[i + 2 + pastnum] + '3', (round((1 + optomistic[i]) * Revenues[pastnum + i], 2)), money)
        Revenues.append((1 + optomistic[i]) * Revenues[pastnum + i])
        ws2.write(ALPHA[i + 1] + '19', round((1 + optomistic[i]) * Revenues[pastnum + i], 2), money)
    "REVENUE PROJECTION NUMBERS"
    for k in range(0, futurenum, 1):
        ws2.write(ALPHA[k + 1] + '20', round((1 + realistic[k]) * Revenues[pastnum + k], 2), money)
        ws2.write(ALPHA[k + 1] + '21', round((1 + pessimistic[k]) * Revenues[pastnum + k], 2), money)
    "THE REST"
    for i in range(0, futurenum, 1):
        ws2.write(ALPHA[i + 2 + pastnum] + '5', round((1 + opcostrt) * Opcosts[pastnum + i], 2), money)
        Opcosts.append((1 + opcostrt) * Opcosts[pastnum + i])
        ws2.write(ALPHA[i + 2 + pastnum] + '6', round((1 + taxrt) * Taxes[pastnum + i], 2), money)
        Taxes.append((1 + taxrt) * Taxes[pastnum + i])
        ws2.write(ALPHA[i + 2 + pastnum] + '7', round((1 + netinvrt) * Netinvests[pastnum + i], 2), money)
        Netinvests.append((1 + netinvrt) * Netinvests[pastnum + i])
        ws2.write(ALPHA[i + 2 + pastnum] + '8', round((1 + chngcaprt) * Workcaps[pastnum + i], 2), money)
        Workcaps.append((1 + chngcaprt) * Workcaps[pastnum + i])
    "CASHFLOWS"
    Cashflow[:] = []
    for y in range(0, spacing + 1, 1):
        ws2.write(ALPHA[y + 1] + '10', round(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]), 2),
                  money)
        Cashflow.append(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]))
    "DEBT & MULTIPLIER"
    ws2.write(ALPHA[2 + pastnum] + '11', Debt, money)
    ws2.write('B13', '-')
    "TERMINAL, ENTERPRISE, FAIR VAL OF EQUITY"
    Fincf = Cashflow[spacing]
    Termval = Fincf * ((1 + revgrowth) / (WACC - revgrowth))
    Entvalue = -(pv(interest, futurenum, 0, Termval))
    Fairval = Entvalue - Debt
    ws2.write('B14', Termval, money)
    ws2.write('B15', Entvalue, money)
    ws2.write('B16', Fairval, yelmon)
    "_______________WS3________________"
    "SHEET SETUP"
    ws3.write('A3', 'Revenues', bold)
    ws3.write('A5', 'Operating Costs', bold)
    ws3.write('A6', 'Taxes', bold)
    ws3.write('A7', 'Net Investments', bold)
    ws3.write('A8', 'Change in Working Capitol', bold)
    ws3.write('A10', 'Cashflow', bold)
    ws3.write('A11', '(Debt)', bold)
    ws3.write('A13', 'Multiplier', bold)
    ws3.write('A14', 'Terminal Value', bold)
    ws3.write('A15', 'Enterprise Value', bold)
    ws3.write('A16', 'Fair Value of Equity', bold)
    ws3.write('A18', 'Revenues', bold)
    ws3.write('A19', 'Optomistic', bold)
    ws3.write('A20', 'Realistic', bold)
    ws3.write('A21', 'Pessimistic', bold)
    count = 0
    for i in range(0, pastnum + 1, 1):
        count = count + 1
        ws3.write(ALPHA[count] + '3', Revenues[count - 1], money)
        ws3.write(ALPHA[count] + '5', Opcosts[count - 1], money)
        ws3.write(ALPHA[count] + '6', Taxes[count - 1], money)
        ws3.write(ALPHA[count] + '7', Netinvests[count - 1], money)
        ws3.write(ALPHA[count] + '8', Workcaps[count - 1], money)
    for i in range(0, pastnum + futurenum + 1, 1):
        ws3.write(ALPHA[i + 1] + '2', years[i], bold)
    ws3.write(ALPHA[spacing + 3] + '2', 'Cost Of Equity', bold)
    ws3.write(ALPHA[spacing + 3] + '3', 'Cost of Debt', bold)
    ws3.write(ALPHA[spacing + 3] + '4', 'beta', bold)
    ws3.write(ALPHA[spacing + 3] + '5', 'Rf', bold)
    ws3.write(ALPHA[spacing + 3] + '6', 'Rm', bold)
    ws3.write(ALPHA[spacing + 3] + '7', 'Re', bold)
    ws3.write(ALPHA[spacing + 3] + '8', 'E/V', bold)
    ws3.write(ALPHA[spacing + 3] + '9', 'D/V', bold)
    ws3.write(ALPHA[spacing + 3] + '10', 'Corp Tax Rate', bold)
    ws3.write(ALPHA[spacing + 3] + '11', 'WACC', bold)
    ws3.write(ALPHA[spacing + 4] + '15', compname, bold)
    "EQUITY VALUES"
    ws3.write(ALPHA[spacing + 4] + '2', Costofeq)
    ws3.write(ALPHA[spacing + 4] + '3', Costofde)
    ws3.write(ALPHA[spacing + 4] + '4', beta)
    ws3.write(ALPHA[spacing + 4] + '5', Rf)
    ws3.write(ALPHA[spacing + 4] + '6', Rm)
    ws3.write(ALPHA[spacing + 4] + '7', Re)
    ws3.write(ALPHA[spacing + 4] + '8', EV)
    ws3.write(ALPHA[spacing + 4] + '9', DV)
    ws3.write(ALPHA[spacing + 4] + '10', Corptaxrate)
    ws3.write(ALPHA[spacing + 4] + '11', WACC)
    "REVENUES"
    Revenues[:] = []
    for i in OrigRevenues:
        Revenues.append(i)
    for i in range(0, futurenum, 1):
        ws3.write(ALPHA[i + 2 + pastnum] + '3', (round((1 + pessimistic[i]) * Revenues[pastnum + i], 2)), money)
        Revenues.append((1 + pessimistic[i]) * Revenues[pastnum + i])
        ws3.write(ALPHA[i + 1] + '21', round((1 + pessimistic[i]) * Revenues[pastnum + i], 2), money)
    "REVENUE PROJECTION NUMBERS"
    for k in range(0, futurenum, 1):
        ws3.write(ALPHA[k + 1] + '20', round((1 + realistic[k]) * Revenues[pastnum + k], 2), money)
        ws3.write(ALPHA[k + 1] + '19', round((1 + optomistic[k]) * Revenues[pastnum + k], 2), money)
    "THE REST"
    for i in range(0, futurenum, 1):
        ws3.write(ALPHA[i + 2 + pastnum] + '5', round((1 + opcostrt) * Opcosts[pastnum + i], 2), money)
        Opcosts.append((1 + opcostrt) * Opcosts[pastnum + i])
        ws3.write(ALPHA[i + 2 + pastnum] + '6', round((1 + taxrt) * Taxes[pastnum + i], 2), money)
        Taxes.append((1 + taxrt) * Taxes[pastnum + i])
        ws3.write(ALPHA[i + 2 + pastnum] + '7', round((1 + netinvrt) * Netinvests[pastnum + i], 2), money)
        Netinvests.append((1 + netinvrt) * Netinvests[pastnum + i])
        ws3.write(ALPHA[i + 2 + pastnum] + '8', round((1 + chngcaprt) * Workcaps[pastnum + i], 2), money)
        Workcaps.append((1 + chngcaprt) * Workcaps[pastnum + i])
    "CASHFLOWS"
    Cashflow[:] = []
    for y in range(0, spacing + 1, 1):
        ws3.write(ALPHA[y + 1] + '10', round(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]), 2),
                  money)
        Cashflow.append(Revenues[y] - (Opcosts[y] + Taxes[y] + Netinvests[y] + Workcaps[y]))
    "DEBT & MULTIPLIER"
    ws3.write(ALPHA[2 + pastnum] + '11', Debt, money)
    ws3.write('B13', '-')
    "TERMINAL, ENTERPRISE, FAIR VAL OF EQUITY"
    Fincf = Cashflow[spacing]
    Termval = Fincf * ((1 + revgrowth) / (WACC - revgrowth))
    Entvalue = -(pv(interest, futurenum, 0, Termval))
    Fairval = Entvalue - Debt
    ws3.write('B14', Termval, money)
    ws3.write('B15', Entvalue, money)
    ws3.write('B16', Fairval, yelmon)
else:
    print('Sorry, answer not recognised...')
"CHART"
chart.add_series({'values': '=Realistic!$B$10:'+ALPHA[spacing+1]+'$10',
                  'categories': '=Realistic!$B$2:'+ALPHA[spacing+1]+'$2',
                  'data_labels': {
                      'value': False,
                      'font': {'name': 'Totals per period'}
                  },
})
chart.set_y_axis({
    'name': 'Cashflow'
})
chart.set_x_axis({
    'name': 'Cashflow in Periods',
})
ws.insert_chart(ALPHA[spacing + 5] + '2', chart)
"CHART 2"
chart2.add_series({'values': '=Optomistic!$B$10:'+ALPHA[spacing+1]+'$10',
                  'categories': '=Optomistic!$B$2:'+ALPHA[spacing+1]+'$2',
                  'data_labels': {
                      'value': False,
                      'font': {'name': 'Totals per period'}
                  },
})
chart2.set_y_axis({
    'name': 'Cashflow'
})
chart2.set_x_axis({
    'name': 'Cashflow in Periods',
})
ws2.insert_chart(ALPHA[spacing + 5] + '2', chart2)
"CHART 3"
chart3.add_series({'values': '=Pessimistic!$B$10:'+ALPHA[spacing+1]+'$10',
                  'categories': '=Pessimistic!$B$2:'+ALPHA[spacing+1]+'$2',
                  'data_labels': {
                      'value': False,
                      'font': {'name': 'Totals per period'}
                  },
})
chart3.set_y_axis({
    'name': 'Cashflow'
})
chart3.set_x_axis({
    'name': 'Cashflow in Periods',
})
ws3.insert_chart(ALPHA[spacing + 5] + '2', chart3)
wb.close()

"""Re Work the cost bands - probably something to do with where Im appending my new numbers"""
"""Actually thats the only major thing - good work"""

sendtext.sendtext('I am finished.')