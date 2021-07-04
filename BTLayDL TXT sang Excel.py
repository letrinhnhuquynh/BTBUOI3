from openpyxl import load_workbook
import io

def read_file_list(file_name):
    f = io.open(file_name, 'r', encoding='utf-8')
    ndung = f.read()
    f.close()
    return ndung.split('\n')

def update_cell(file_path,sheetname,cell_name,new_value):
    wb = load_workbook(filename = file_path)
    wb[sheetname][cell_name].value = new_value
    wb.close()
    wb.save(file_path)

if __name__ == "__main__":
    ds1 = "abc.txt"
    ds2 = "abc2.txt"
    N1 = read_file_list(ds1)
    N2 = read_file_list(ds2)
    file_path = 'test.xlsx'
    sheetname = 'quynh'
    update_cell(file_path, sheetname, 'A1', 'Tên')
    update_cell(file_path, sheetname, 'B1', 'Tuổi')

    for i in range(0, len(N1)):
        ten = N1[i]
        tuoi = N2[i]
        cell_name = 'A%s' % (i + 2)
        update_cell(file_path, sheetname, cell_name, ten)
        cell_name = 'B%s' % (i + 2)
        update_cell(file_path, sheetname, cell_name, tuoi)


