import math
import openpyxl

print("Enter the Primary Voltage Vp in volt (V)")
Vp = float(input())

print("Enter the Secondary Voltage Vs in volt (V)")
Vs = float(input())

print("Enter the Secondary Current Is in ampere (A)")
Is = float(input())

print("Enter the value of K through the table")
K = float(input())

print("Enter the value of frequency f in Hz")
f = float(input())

print("Enter the value of Magnetic flux density Bmax in wb/m²")
Bmax = float(input())

print("Enter the value of current density J in A/m²")
J = float(input())

print("Enter the value of paper width Wp in mm")
Wp = float(input())

Ep = Vp
Es = Vs
Ip = (Vs * Is) / Vp
Ss = ((Vs * Is) / 1000)
print("Power of secondary side (Ss) is", Ss, "KVA")

Q = math.sqrt(Ss)
Et = round(K * Q, 3)
print("Emf per turn of the coil Et is", Et)
Np = int(Vp / Et)
print("Number of primary turns per phase Np is", Np, "turns")
Ns = int((Vs * Np) / Vp)
print("Number of secondary turns per phase Ns is", Ns, "turns")
Aw = int(Ep / (4.44 * f * Bmax * Np * 0.98 * pow(10, -6)))
print("Calculating the area of the bobbin Aw is", Aw, "mm²")
Wl = int(math.sqrt(Aw))
print("Area of one side is", Wl, "mm")

A1 = round(Ip / J, 3)
print("The cross-section of primary conductor is: ", A1)
A2 = round(Is / J, 3)
print("The cross-section of secondary conductor is: ", A2)

# Load the Excel workbook for Sankey Number
sankey_workbook = openpyxl.load_workbook('sankey.xlsx')
sankey_worksheet = sankey_workbook['Sheet1']

# Get the input value for Wl
input_value1 = Wl

closest_rows = []
closest_abs_diff = float('inf')

for row in sankey_worksheet.iter_rows(min_row=2, max_row=sankey_worksheet.max_row, min_col=2, max_col=2):
    cell_value = row[0].value
    if cell_value is not None and isinstance(cell_value, (int, float)):
        abs_diff = abs(cell_value - input_value1)
        if abs_diff <= closest_abs_diff:
            if abs_diff < closest_abs_diff:
                closest_rows.clear()
            closest_abs_diff = abs_diff
            closest_rows.append([cell.value for cell in sankey_worksheet[row[0].row]])

if closest_rows:
    print("Rows of the closest cells:")
    for row_values in closest_rows:
        print(*row_values, sep=', ')

WH = float(input("Enter the value of WH: "))

closest_row = None
closest_abs_diff = float('inf')

for row_values in closest_rows:
    cell_value = row_values[2]
    abs_diff = abs(cell_value - WH)
    if abs_diff < closest_abs_diff:
        closest_abs_diff = abs_diff
        closest_row = row_values

if closest_row is not None:
    print("Row of the closest cell for WH:")
    print(*closest_row, sep=', ')

Wl = row_values[1]
Sankey_num = row_values[4]

print(f'The available size of Wl is: {Wl}')
print(f'The Sankey Number is: {Sankey_num}')

# Load the Excel workbook for Conductor
conductor_workbook = openpyxl.load_workbook('conductor.xlsx')
conductor_worksheet = conductor_workbook['Sheet1']

# Get the input value for A1
input_value1 = A1

closest_rows = []
closest_abs_diff = float('inf')

for row in conductor_worksheet.iter_rows(min_row=2, max_row=conductor_worksheet.max_row, min_col=2, max_col=2):
    cell_value = row[0].value
    if cell_value is not None and isinstance(cell_value, (int, float)):
        abs_diff = abs(cell_value - input_value1)
        if abs_diff <= closest_abs_diff:
            if abs_diff < closest_abs_diff:
                closest_rows.clear()
            closest_abs_diff = abs_diff
            closest_rows.append([cell.value for cell in conductor_worksheet[row[0].row]])

if closest_rows:
    print("Rows of the closest cells:")
    for row_values in closest_rows:
        print(*row_values, sep=', ')

SWG = row_values[3]

print(f'The nearest value of A1 is {row_values[1]}')
print(f'The diameter of A1 is {row_values[2]}')
print(f'The available size of SWG is: {SWG}')

# Get the input value for A2
input_value2 = A2

closest_rows = []
closest_abs_diff = float('inf')

for row in conductor_worksheet.iter_rows(min_row=2, max_row=conductor_worksheet.max_row, min_col=2, max_col=2):
    cell_value = row[0].value
    if cell_value is not None and isinstance(cell_value, (int, float)):
        abs_diff = abs(cell_value - input_value2)
        if abs_diff <= closest_abs_diff:
            if abs_diff < closest_abs_diff:
                closest_rows.clear()
            closest_abs_diff = abs_diff
            closest_rows.append([cell.value for cell in conductor_worksheet[row[0].row]])

if closest_rows:
    print("Rows of the closest cells:")
    for row_values in closest_rows:
        print(*row_values, sep=', ')

SWG = row_values[3]

print(f'The nearest value of A2 is {row_values[1]}')
print(f'The diameter of A2 is {row_values[2]}')
print(f'The available size of SWG is: {SWG}')