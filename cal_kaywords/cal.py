import sys
print("start procedure...")
import openpyxl

'''
Requirement:
    1. filename
    2. to-be-analyzed sheet should named 'tb1'
'''


#read xls
#TODO: handle file extention error
print("reading xls'x'.....")
FILE_NAME = "C:\\Users\wkroh\\Documents\\vscode_github\miniUtil\\miniUtil\\cal_kaywords\\full.xlsx"

RESULT_SHEET_NAMES=("Result_Author_Keyword", "Result_Keyword_Plus")

try:
    wb = openpyxl.load_workbook(FILE_NAME)
    print("reading workbook successfully")
except Exception as e:
    print(f"[Error] Your FILE_NAME '{FILE_NAME}' does not exist. ErrCode:{repr(e)}")
    sys.exit(1)  
try:
    wb.save(FILE_NAME)
except Exception as e:
    print(f"[Error] You probabaly opened {FILE_NAME} by MS Excel. ErrCode:{repr(e)}")
    sys.exit(1)

for rst in RESULT_SHEET_NAMES:
    if rst in wb.sheetnames:
        print('[Warn] previous result will be removed')
        wb.remove(wb[rst])
        wb.save(FILE_NAME)

tb = wb['tb1']

article_title_idx=-1
author_keywords_idx=-1
keywords_plus_idx=-1
g_kw_feq = {}
g_kwp_feq = {}

for row_idx, row in enumerate(tb.iter_rows(values_only=True), start=1):
    if row_idx == 1:
        for idx, cell in enumerate(row):
            if cell == 'Article Title':
                article_title_idx = idx
            elif cell == 'Author Keywords':
                author_keywords_idx = idx
            elif cell == 'Keywords Plus':
                keywords_plus_idx = idx
        print(f"article_title_idx: {article_title_idx}, author_keywords_idx:{author_keywords_idx}, keywords_plus_idx:{keywords_plus_idx}")
        continue
    try:
        #print(row)
        kw_cell = row[author_keywords_idx]
        kw_cell2 = row[keywords_plus_idx]
        if not kw_cell:
            raise Exception(f"empty 'Author Keywords' at row {row_idx}")
        if not kw_cell2:
            raise Exception(f"empty 'Keywords Plus' at row {row_idx}")

        kw_list = [k.strip() for k in kw_cell.split(';')]
        for kw in kw_list:
            lkw = kw.lower()
            if lkw not in g_kw_feq.keys():
                g_kw_feq[lkw] = 1
            else:
                g_kw_feq[lkw] += 1

        kw_list = [k.strip() for k in kw_cell2.split(';')]
        for kw in kw_list:
            if kw not in g_kwp_feq.keys():
                g_kwp_feq[kw] = 1
            else:
                g_kwp_feq[kw] += 1
    except Exception as e:
        print(f'[WARN] row {row_idx} has unexpected format. ErrCode:{repr(e)}')

result_sheet = wb.create_sheet("Result_Author_Keyword")
for pair in g_kw_feq.items():
    result_sheet.append(pair)

result_sheet = wb.create_sheet("Result_Keyword_Plus")
for pair in g_kwp_feq.items():
    result_sheet.append(pair)
wb.save(FILE_NAME)